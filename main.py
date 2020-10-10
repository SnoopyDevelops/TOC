# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import os
# from sys import argv
from argparse import ArgumentParser
from math import ceil

import docx2pdf
import fitz
# import img2pdf
from PIL import Image
from PyPDF2 import PdfFileReader, PdfFileWriter  # , PdfFileMerger


# from fpdf import FPDF


def convert_doc_to_pdf(doc_file):
    # Use a breakpoint in the code line below to debug your script.
    # Press Ctrl+F8 to toggle the breakpoint.
    filename, _ = os.path.splitext(doc_file)
    pdf_file = filename + '.pdf'
    docx2pdf.convert(doc_file, pdf_file)

    return pdf_file


# def convert_image_to_pdf(image_file):
#     filename, _ = os.path.splitext(image_file)
#     image = Image.open(image_file)
#     if image.mode == 'RGBA':
#         bg = Image.new('RGB', image.size, (255, 255, 255))
#         bg.paste(image, (0, 0), image)
#         bg.save('alpha_removed.png', 'PNG')
#         image_file = 'alpha_removed.png'
#
#     pdf_file = filename + '.pdf'
#     with open(pdf_file, 'wb') as file:
#         file.write(img2pdf.convert(image_file))
#         if image_file == 'alpha_removed.png':
#             os.remove(image_file)
#
#     return pdf_file


def text_to_sequence(text, font):
    max_lens = {16: 70, 14: 80, 12: 90}
    max_len = max_lens[font]
    words = text.split()
    sequence = []
    line = ''
    while len(words) > 0:
        word = words.pop(0)
        if len(line) + len(word) < max_len:
            line = ' '.join([line, word])
        else:
            sequence.append(line)
            line = word
    sequence.append(line)
    return sequence


def main():
    # if len(argv) == 2:
    #     folder = argv[1]
    # else:
    #     folder = '.\\PDF Working Data\\Sample 1'

    parser = ArgumentParser()
    parser.add_argument('--folder', type=str, default='.\\PDF Working Data\\Sample 2')
    parser.add_argument('--logo', type=str, default='logo.png')
    parser.add_argument('--title', type=str, default='Complete Subscription Package')
    parser.add_argument('--sub_title', type=str, default='As Of October 9, 2020')
    parser.add_argument('--comment', type=str,
                        default='This is a paragraph which should appear after the sub-title, '
                                'but before the start of the links/items in the table of contents'
                        )

    args = parser.parse_args()
    folder = args.folder

    # merger = PdfFileMerger()
    # merge_list = []
    merged_file = os.path.join(folder, 'merged.pdf')
    converted_files, image_files = [], []
    bookmarks = []
    page = 0
    docs = fitz.open()

    for file in os.listdir(folder):
        filename, extension = os.path.splitext(file)
        file = os.path.join(folder, file)
        if extension in ['.png', '.jpg']:
            # file = convert_image_to_pdf(file)
            # converted_files.append(file)
            image_files.append(file)
        else:
            if extension in ['.doc', '.docx']:
                file = convert_doc_to_pdf(file)
                converted_files.append(file)
            # merge_list.append(file)
            # merger.append(file, import_bookmarks=False)
            # merger.addBookmark(title=filename, pagenum=page)
            # page += PdfFileReader(file).numPages
            with fitz.open(file) as doc:
                docs.insertPDF(doc)
                bookmarks.append((filename, page + 1))
                page += doc.pageCount

    # merger.write(merged_file)
    # merger.close()

    # docs = fitz.open(merged_file)
    width, height = 612, 792
    for i, file in enumerate(image_files):
        image = Image.open(file)
        image_width, image_height = image.size

        if image_height < image_width * 1.3:
            x0, x1 = 115, 497
            h = int(image_height * 382 / image_width)
            y0 = int(396 - h / 2)
            y1 = 792 - y0
        else:
            y0, y1 = 149, 643
            w = int(image_width * 494 / image_height)
            x0 = int(306 - w / 2)
            x1 = 612 - x0

        img_rect = fitz.Rect(x0, y0, x1, y1)
        docs.newPage(-1, width=width, height=height)
        docs[page].insertImage(img_rect, filename=file)

        filename, _ = os.path.splitext(file)
        filename = filename.split('\\')[-1]
        bookmarks.append((filename, page + 1))
        page += 1

    docs.newPage(0, width=width, height=height)

    # insert logo image
    logo = Image.open(args.logo)
    logo_width, logo_height = logo.size
    x1, y0, y1 = 576, 36, 108
    x0 = x1 - 72 * logo_width / logo_height
    img_rect = fitz.Rect(x0, y0, x1, y1)
    docs[0].insertImage(img_rect, filename=args.logo)

    # insert title
    y = 108
    title_font, sub_title_font, comment_font = 16, 14, 12
    title = text_to_sequence(text=args.title, font=title_font)
    line_number = docs[0].insertText(point=fitz.Point(108, y), text=title, fontsize=title_font)

    # insert subtitle
    y += ceil(line_number / 2) * 36
    sub_title = text_to_sequence(text=args.sub_title, font=sub_title_font)
    line_number = docs[0].insertText(point=fitz.Point(144, y), text=sub_title, fontsize=sub_title_font)

    # insert comment
    y += ceil(line_number / 2) * 36
    comment = text_to_sequence(text=args.comment, font=comment_font)
    line_number = docs[0].insertText(point=fitz.Point(72, y), text=comment, fontsize=comment_font)

    # x, y = 0, 36
    y += ceil(line_number / 2) * 36
    toc_y = y
    for bookmark in bookmarks:
        section, page = bookmark
        # text = '-' * 140
        # text = ' '.join([section, '-' * (100 - len(section)), str(page + 1)])
        docs[0].insertText(point=fitz.Point(36, y), text=section)
        docs[0].insertText(point=fitz.Point(540, y), text=str(page + 1))
        y += 36

    docs.save(merged_file)
    docs.close()

    for file in converted_files:
        os.remove(file)

    reader, writer = PdfFileReader(merged_file), PdfFileWriter()
    for page in range(reader.numPages):
        page = reader.getPage(page)
        page.scaleTo(width=width, height=height)
        writer.addPage(page)

    with open("scaled.pdf", "wb") as f:
        writer.write(f)

    os.remove(merged_file)

    # open_file = open(merged_file, 'rb')
    open_file = open('scaled.pdf', 'rb')
    pdf_file = PdfFileReader(open_file)

    writer = PdfFileWriter()
    for i in range(pdf_file.getNumPages()):
        writer.addPage(pdf_file.getPage(i))

    xll, xur, yur = 36, 576, 792 - toc_y
    yll = yur + 18
    for bookmark in bookmarks:
        _, page = bookmark
        writer.addLink(pagenum=0, pagedest=page, rect=[xll, yll, xur, yur])
        yll -= 36
        yur -= 36

    filename = folder.split('\\')[-1]
    with open(filename + '.pdf', 'wb') as write_file:
        writer.write(write_file)

    open_file.close()
    os.remove('scaled.pdf')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()
