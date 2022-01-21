# - * - coding:utf-8 - * -
#  作者：Elias Cheung
#  编写时间：2022/1/20  14:42
import math
from xlwt import *
import xlwings as xw
from PIL import Image
from typing import Tuple


def resize_image(original_pic_path: str, resized_pic_path: str, max_resize_width: int = 256) -> Tuple[str, tuple]:
    img = Image.open(original_pic_path)
    img_width, img_height = img.size

    resize_width, resize_height = max_resize_width, math.ceil(img_height * max_resize_width / img_width)
    # print(resize_width, resize_height)
    resized_image = img.resize((resize_width, resize_height))
    resized_image.save(resized_pic_path)
    return resized_pic_path, (resize_width, resize_height)


def pixel_colour(image_path: str, image_size: tuple) -> list:
    img = Image.open(image_path)
    img = img.convert('RGB')
    colour_list = img.load()

    img_width, img_height = image_size
    pixel_coordinates = [(i, j) for i in range(img_width) for j in range(img_height)]

    coordinates_and_colour_to_be_returned = [(colour_list[i, j], (i, j)) for i, j in pixel_coordinates]

    return coordinates_and_colour_to_be_returned


def create_excel(size: tuple, cell_width: int=2) -> str:
    cell_width = 256 * cell_width
    _book = Workbook(encoding="utf - 8")
    sheet = _book.add_sheet("picture", cell_overwrite_ok=True)
    for i in range(size[0]):
        first_col = sheet.col(i)
        first_col.width = cell_width
    book_name = "picture.xls"
    _book.save(book_name)
    return book_name


def draw_excel(excel: str, colour_and_cord: list):
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(excel)
    sht = wb.sheets[0]
    print(sht)
    for cell_info in colour_and_cord:
        r, g, b = cell_info[0]
        x, y = cell_info[1]
        try:
            sht.range((y + 1, x + 1), (y + 1, x + 1)).color = (r, g, b)
        except ValueError as e:
            print(f"cell index out of range: {e}")
            continue
    wb.save("processed.xlsx")


if __name__ in "__main__":
    resized = resize_image(original_pic_path="IMG_4531.JPG", resized_pic_path="IMG_4531.resize.png")
    print(resized)
    colours = pixel_colour(resized[0], resized[1])
    book = create_excel(resized[1])
    draw_excel(book, colours)
