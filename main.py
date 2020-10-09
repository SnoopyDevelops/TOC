# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import os
from sys import argv

import docx2pdf
import img2pdf
from PIL import Image
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
from fpdf import FPDF


def convert_doc_to_pdf(doc_file):
    # Use a breakpoint in the code line below to debug your script.
    # Press Ctrl+F8 to toggle the breakpoint.
    filename, _ = os.path.splitext(doc_file)
    # print(filename)
    pdf_file = filename + '.pdf'
    docx2pdf.convert(doc_file, pdf_file)
    # os.remove(doc_file)

    return pdf_file


def convert_image_to_pdf(image_file):
    # filename = '.'.join(image_file.split('.')[:-1])
    filename, _ = os.path.splitext(image_file)
    image = Image.open(image_file)
    if image.mode == 'RGBA':
        bg = Image.new('RGB', image.size, (255, 255, 255))
        bg.paste(image, (0, 0), image)
        bg.save('alpha_removed.png', 'PNG')
        # os.remove(image_file)
        image_file = 'alpha_removed.png'

    pdf_file = filename + '.pdf'
    with open(pdf_file, 'wb') as file:
        file.write(img2pdf.convert(image_file))
        if image_file == 'alpha_removed.png':
            os.remove(image_file)

    return pdf_file


def main():
    if len(argv) == 2:
        folder = argv[1]
    else:
        folder = '.\\PDF Working Data\\Sample 1'
    merged_file = os.path.join(folder, 'merged.pdf')

    if os.path.exists(merged_file):
        os.remove(merged_file)

    converted_files, image_files = [], []
    merger = PdfFileMerger()
    bookmarks, merge_list = [], []
    page = 0

    for file in os.listdir(folder):
        # print(file)
        filename, extension = os.path.splitext(file)
        file = os.path.join(folder, file)
        if extension in ['.doc', '.docx']:
            file = convert_doc_to_pdf(file)
            converted_files.append(file)
            merge_list.append(file)
            # merger.addBookmark(title=filename, pagenum=page)
            bookmarks.append((filename, page + 1))
            page += PdfFileReader(file).numPages
        elif extension in ['.png', '.jpg']:
            file = convert_image_to_pdf(file)
            converted_files.append(file)
            image_files.append(file)
        else:
            merge_list.append(file)
            # merger.addBookmark(title=filename, pagenum=page)
            bookmarks.append((filename, page + 1))
            page += PdfFileReader(file).numPages

    for file in image_files:
        merge_list.append(file)
        filename, _ = os.path.splitext(file)
        filename = filename.split('\\')[-1]
        bookmarks.append((filename, page + 1))
        page += 1

    toc = FPDF(unit='pt', format='letter')
    toc.add_page()
    toc.set_font('Arial')
    x, y = 0, 36
    for bookmark in bookmarks:
        section, page = bookmark
        toc.cell(w=x, h=y, txt=section, border=1, ln=1)
    toc.output('toc.pdf', 'F')

    merger.append('toc.pdf')
    for file in merge_list:
        merger.append(file, import_bookmarks=False)

    merger.write(merged_file)
    merger.close()

    for file in converted_files:
        os.remove(file)

    os.remove('toc.pdf')

    newWidth, newHeight = 612, 792

    reader, writer = PdfFileReader(merged_file), PdfFileWriter()
    for page in range(reader.numPages):
        page = reader.getPage(page)

        page.scaleTo(newWidth, newHeight)

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

    xll, yll, xur, yur = 29, 763, 584, 727

    for bookmark in bookmarks:
        section, page = bookmark
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
