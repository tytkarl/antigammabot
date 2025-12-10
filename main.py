import sys
import zipfile
import xml.etree.ElementTree as ET
import io
import os
from asyncio import sleep
from PIL import Image
import fitz

from dotenv import load_dotenv
load_dotenv()

from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    filters,
    ContextTypes,
    CommandHandler
)
from telegram.request import HTTPXRequest

BOT_TOKEN = os.getenv("BOT_TOKEN") #by Slava Bebrou

def clean_pptx(input_path, output_path):
    ns = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
    }

    watermark_sizes = [(921, 220), (1842, 440)]
    removed_images = set()

    modified_xml = {}

    with zipfile.ZipFile(input_path, 'r') as zin:
        filelist = zin.namelist()

        slides  = [f for f in filelist if f.startswith("ppt/slides/slide") and f.endswith(".xml")]
        layouts = [f for f in filelist if f.startswith("ppt/slideLayouts/slideLayout") and f.endswith(".xml")]
        masters = [f for f in filelist if f.startswith("ppt/slideMasters/slideMaster") and f.endswith(".xml")]

        def process_file(xml_path, base_dir):
            filename = xml_path.split('/')[-1]
            rels_dir  = base_dir + "/_rels/"
            rels_path = rels_dir + filename + ".rels"

            image_rels = {}

            # читаем rels
            if rels_path in filelist:
                root_rels = ET.fromstring(zin.read(rels_path))

                for rel in list(root_rels):
                    if not rel.tag.endswith('Relationship'):
                        continue

                    rid = rel.get('Id')
                    tgt = rel.get('Target', '')
                    rType = rel.get('Type', '')

                    if 'image' in rType:
                        image_rels[rid] = tgt

                modified_xml[rels_path] = ET.tostring(root_rels, encoding='utf-8', xml_declaration=True)
            else:
                root_rels = None

            # читаем xml
            root = ET.fromstring(zin.read(xml_path))
            spTree = root.find('.//p:spTree', ns)
            if spTree is None:
                return
            
            pic_ids = []

            # поиск картинок
            for pic in list(spTree.findall('p:pic', ns)):
                blip = pic.find('.//a:blip', ns)
                if blip is None:
                    continue

                embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if embed not in image_rels:
                    continue

                target_file = image_rels[embed]

                if target_file.startswith("../"):
                    target_file = target_file[3:]
                if not target_file.startswith("ppt/"):
                    target_file = "ppt/" + target_file

                if target_file not in filelist:
                    continue

                try:
                    img = Image.open(io.BytesIO(zin.read(target_file)))
                except:
                    continue

                # watermark gamma
                if (img.width, img.height) in watermark_sizes:
                    spTree.remove(pic)
                    pic_ids.append(embed)
                    removed_images.add(target_file)

            # чистим rels
            if root_rels is not None:
                for rel in list(root_rels):
                    if rel.get('Id') in pic_ids:
                        root_rels.remove(rel)

                modified_xml[rels_path] = ET.tostring(root_rels, encoding='utf-8', xml_declaration=True)

            # сохраняем изменённый xml
            modified_xml[xml_path] = ET.tostring(root, encoding='utf-8', xml_declaration=True)

        # обрабатываем все файлы
        for slide in slides:
            process_file(slide, "ppt/slides")
        for layout in layouts:
            process_file(layout, "ppt/slideLayouts")
        for master in masters:
            process_file(master, "ppt/slideMasters")

        # создаём новый pptx
        with zipfile.ZipFile(output_path, 'w') as zout:
            for item in filelist:
                if item in removed_images:
                    continue
                zout.writestr(item, modified_xml.get(item, zin.read(item)))

# PDF cleaner
def clean_pdf(input_path, output_path):
    import fitz
    doc = fitz.open(input_path)

    SAMPLE_OFFSETS = [
        (0.5, 0.5),   # центр
        (0.1, 0.5),   # слева по центру
        (0.9, 0.5),   # справа по центру
        (0.5, 0.2),   # сверху по центру
        (0.5, 0.8),   # снизу по центру
    ]

    for page in doc:
        W, H = page.rect.width, page.rect.height

        # прямоугольник водяного знака
        rect = fitz.Rect(W - 2000, H - 37, W, H)

        colors = []

        for ox, oy in SAMPLE_OFFSETS:
            sx = rect.x0 + rect.width * ox
            sy = rect.y0 + rect.height * oy

            # маленькая выборка 1×1
            clip_rect = fitz.Rect(sx, sy, sx + 1, sy + 1)

            try:
                pix = page.get_pixmap(clip=clip_rect)
                r, g, b = pix.samples[:3]
                colors.append((r, g, b))
            except:
                pass

        # если не удалось собрать ни одной точки
        if not colors:
            avg_color = (1, 1, 1)
        else:
            # усреднение RGB
            r = sum(c[0] for c in colors) / (255 * len(colors))
            g = sum(c[1] for c in colors) / (255 * len(colors))
            b = sum(c[2] for c in colors) / (255 * len(colors))
            avg_color = (r, g, b)

        # закрашиваем область
        page.draw_rect(rect, fill=avg_color, color=avg_color, overlay=True)

    doc.save(output_path)
    doc.close()


# ЭЭЭ ТЕЛЕГРАМ БОТ ЕЖЖЕ БЛЯ
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Отправь PDF или PPTX — я очищу watermark Gamma.\n ФАЙЛ PPTX ДО 20 МБ, ФАЙД PDF ТОЖЕ!")

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    fname = doc.file_name.lower()

    os.makedirs("tmp", exist_ok=True)
    os.makedirs("out", exist_ok=True)

    in_path = f"tmp/{fname}"
    out_path = f"out/cleaned_{fname}"

    file = await doc.get_file()
    await file.download_to_drive(in_path)

    if fname.endswith(".pptx"):
        clean_pptx(in_path, out_path)
    elif fname.endswith(".pdf"):
        clean_pdf(in_path, out_path)
    else:
        await update.message.reply_text("Баклан тупой я же сказал только PDF ИЛИ PPTX.")
        return

    await update.message.reply_document(
        document=open(out_path, "rb"),
        filename=(fname)
    )

    await update.message.reply_text("С любовью, Слава Беброу.")
    await sleep(0,5) 
    await update.message.reply_text("❤️")


def main():
    req = HTTPXRequest(connect_timeout=20, read_timeout=200)
    app = ApplicationBuilder().token(BOT_TOKEN).request(req).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))

    app.run_polling()

if __name__ == "__main__":
    main()
