import sys
import zipfile
import xml.etree.ElementTree as ET
import io
import os
from PIL import Image
import fitz

from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    filters,
    ContextTypes,
    CommandHandler
)
from telegram.request import HTTPXRequest


# ----------------------------- PPTX CLEANER -----------------------------

def clean_pptx(input_path, output_path):
    ns = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
    }

    watermark_sizes = [(921, 220), (1842, 440)]
    removed_images = set()
    images_removed = 0
    links_removed = 0
    modified_xml = {}

    with zipfile.ZipFile(input_path, 'r') as zin:
        filelist = zin.namelist()

        slides  = [f for f in filelist if f.startswith("ppt/slides/slide") and f.endswith(".xml")]
        layouts = [f for f in filelist if f.startswith("ppt/slideLayouts/slideLayout") and f.endswith(".xml")]
        masters = [f for f in filelist if f.startswith("ppt/slideMasters/slideMaster") and f.endswith(".xml")]

        def process_file(xml_path, base_dir):
            nonlocal images_removed, links_removed
            filename = xml_path.split('/')[-1]
            rels_dir  = base_dir + "/_rels/"
            rels_path = rels_dir + filename + ".rels"

            image_rels = {}
            hlink_ids = []

            if rels_path in filelist:
                rel_data = zin.read(rels_path)
                root_rels = ET.fromstring(rel_data)

                for rel in list(root_rels):
                    if not rel.tag.endswith('Relationship'):
                        continue
                    rid = rel.get('Id')
                    target = rel.get('Target', '')
                    rType = rel.get('Type', '')

                    if 'image' in rType:
                        image_rels[rid] = target

                    if 'hyperlink' in rType and 'gamma.app' in target:
                        hlink_ids.append(rid)
                        root_rels.remove(rel)
                        links_removed += 1

                new_rels_xml = ET.tostring(root_rels, encoding='utf-8', xml_declaration=True)
                modified_xml[rels_path] = new_rels_xml
            else:
                root_rels = None

            xml_data = zin.read(xml_path)
            root = ET.fromstring(xml_data)
            spTree = root.find('.//p:spTree', ns)
            if spTree is None:
                return

            keywords = ["gamma", "button", "watermark"]
            to_remove = []

            for sp in spTree.findall('p:sp', ns):
                texts = [t.text for t in sp.findall('.//a:t', ns) if t.text]
                combined = (" ".join(texts)).lower()
                if any(k in combined for k in keywords):
                    to_remove.append(sp)

            for sp in to_remove:
                spTree.remove(sp)

            pic_remove = []
            pic_ids = []
            for pic in spTree.findall('p:pic', ns):
                blip = pic.find('.//a:blip', ns)
                if blip is None:
                    continue

                embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if not embed or embed not in image_rels:
                    continue

                target = image_rels[embed]
                if target.startswith("../"):
                    target_file = target[3:]
                else:
                    target_file = target

                if not target_file.startswith("ppt/"):
                    target_file = "ppt/" + target_file

                if target_file not in filelist:
                    continue

                data = zin.read(target_file)
                try:
                    img = Image.open(io.BytesIO(data))
                except:
                    continue

                size = (img.width, img.height)
                if size in watermark_sizes:
                    pic_remove.append(pic)
                    pic_ids.append(embed)
                    removed_images.add(target_file)

            for pic in pic_remove:
                spTree.remove(pic)
                images_removed += 1

            if root_rels is not None:
                for rel in list(root_rels):
                    if rel.get('Id') in pic_ids:
                        root_rels.remove(rel)

                new_rels_xml = ET.tostring(root_rels, encoding='utf-8', xml_declaration=True)
                modified_xml[rels_path] = new_rels_xml

            for elem in root.iter():
                for child in list(elem):
                    if child.tag.endswith('hlinkClick'):
                        if child.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') in hlink_ids:
                            elem.remove(child)

            new_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            modified_xml[xml_path] = new_xml

        for slide in slides:
            process_file(slide, "ppt/slides")
        for layout in layouts:
            process_file(layout, "ppt/slideLayouts")
        for master in masters:
            process_file(master, "ppt/slideMasters")

        with zipfile.ZipFile(output_path, 'w') as zout:
            for item in filelist:
                if item in removed_images:
                    continue
                if item in modified_xml:
                    zout.writestr(item, modified_xml[item])
                else:
                    zout.writestr(item, zin.read(item))

    print(f"PPTX очищено — удалено {images_removed} watermark-картинок и {links_removed} ссылок.")


# ----------------------------- PDF CLEANER -----------------------------

def clean_pdf(input_path, output_path):
    doc = fitz.open(input_path)

    for page in doc:
        W = page.rect.width
        H = page.rect.height

        rect = fitz.Rect(W - 2000, H - 37, W, H)

        sample_x = max(0, rect.x0 - 10)
        sample_y = max(0, rect.y0 + rect.height / 2)

        try:
            pix = page.get_pixmap(clip=fitz.Rect(sample_x, sample_y, sample_x + 1, sample_y + 1))
            bg = tuple(pix.samples[:3])
            color = (bg[0] / 255, bg[1] / 255, bg[2] / 255)
        except:
            color = (1, 1, 1)

        page.draw_rect(rect, fill=color, color=color, overlay=True)

    doc.save(output_path)
    doc.close()


# ----------------------------- BOT LOGIC -----------------------------

TOKEN = ""   #НУ ТИПА ЭТА ТАКЕН


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Отправь PDF или PPTX — я очищу watermark Gamma."
    )


async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    fname = doc.file_name.lower()

    file = await doc.get_file()

    os.makedirs("tmp", exist_ok=True)
    os.makedirs("out", exist_ok=True)

    in_path = f"tmp/{fname}"
    out_path = f"out/cleaned_{fname}"

    await file.download_to_drive(in_path)

    if fname.endswith(".pptx"):
        clean_pptx(in_path, out_path)
    elif fname.endswith(".pdf"):
        clean_pdf(in_path, out_path)
    else:
        await update.message.reply_text("Поддерживаются только PDF и PPTX.")
        return

    await update.message.reply_document(
        document=open(out_path, "rb"),
        filename=f"cleaned_{fname}"
    )


def main():
    req = HTTPXRequest(connect_timeout=20, read_timeout=200)
    app = ApplicationBuilder().token(TOKEN).request(req).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))

    app.run_polling()


if __name__ == "__main__":
    main()
