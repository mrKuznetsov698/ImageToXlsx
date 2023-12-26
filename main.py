import xlsxwriter.worksheet
from PIL import Image
import xlsxwriter

SQ_SIZE = 4
LINE_SIZE = 0
FILENAME = 'nyan_cat-1.png'


def get_avrg_col(img: Image, x1, y1, wid, hei):
    x2 = x1 + wid
    y2 = y1 + hei
    r = 0
    g = 0
    b = 0
    for x in range(x1, x2):
        for y in range(y1, y2):
            rt, gt, bt = img.getpixel((x, y))
            r += rt
            g += gt
            b += bt
    r = r // (x2 - x1) // (y2 - y1)
    g = g // (x2 - x1) // (y2 - y1)
    b = b // (x2 - x1) // (y2 - y1)
    return r, g, b


def format_color(r, g, b):
    return '#{:02x}{:02x}{:02x}'.format(r, g, b)


def set_cell_color(worksheet: xlsxwriter.worksheet.Worksheet, workbook: xlsxwriter.Workbook, x, y, r, g, b,):
    format = workbook.add_format()
    format.set_bg_color(bg_color=format_color(r, g, b))
    worksheet.write_blank(row=y, col=x, blank=None, cell_format=format)


with xlsxwriter.Workbook('output.xlsx') as workbook:
    # Image
    img = Image.open(FILENAME).convert('RGB')
    WIDTH = img.width
    HEIGHT = img.height
    WT = WIDTH // (SQ_SIZE + LINE_SIZE)
    HT = HEIGHT // (SQ_SIZE + LINE_SIZE)
    # Worksheet
    worksheet = workbook.add_worksheet()
    # Set row & column size
    for x in range(WT):
        for y in range(HT):
            worksheet.set_row(row=y, height=SQ_SIZE)
            worksheet.set_column(first_col=0, last_col=WT, width=SQ_SIZE//4)
    # Main loop
    for x in range(WT):
        for y in range(HT):
            print(x, y)
            print(x * SQ_SIZE, y * SQ_SIZE)
            r, g, b = get_avrg_col(img, (x + 1) * LINE_SIZE + x * SQ_SIZE, (y + 1) * LINE_SIZE + y * SQ_SIZE, SQ_SIZE, SQ_SIZE)
            set_cell_color(worksheet, workbook, x, y, r, g, b)
