# !/usr/bin/python3
# -*- coding: utf-8 -*-
import argparse
import io
import math
import os
import shutil
import subprocess
import xml.etree.ElementTree as ET

from pandas import read_csv
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

class IncoherentDataError(Exception):
    pass

PATH_REPLACE_RULES = [
    (r'C:\Users\Florence\Pictures\2015\Roadtrip\1Best of roadtrip', '/home/sindarus/mapalb2pdf/images/all2/'),
    (r'C:\Users\Florence\Pictures\2015\Roadtrip\J11 - 31 juillet Creuse (Limoges - Oradour sur Glane)', '/home/sindarus/mapalb2pdf/images/all2/'),
    (r'C:\Users\Florence\Pictures\2015\Roadtrip\J8 - 28 juillet Ardèche Gard (kayak -', '/home/sindarus/mapalb2pdf/images/all2/'),
    (r'C:\Users\Florence\Pictures\2015\Roadtrip\J1 - 21 juillet Ardennes - Marne', '/home/sindarus/mapalb2pdf/images/all2/'),
]
PAGE_SIZE = (841.89, 595.27)

FONT_SIZE_CORRECT_FACTOR = 75/100

DO_ROTATE = True
ZOOM_ROTATE = True
CLIP_IMAGES = True
DRAW_BORDER = False


def parse_my_args():
    parser = argparse.ArgumentParser(description='Produces a PDF file from a mapalb file.')
    parser.add_argument('-i', '--input', dest='input_file', required=True,
                        help='Input mapalb file')
    parser.add_argument('-o', '--output', dest='output_file', required=True,
                        help='Output pdf file')
    return parser.parse_args()


def load_mdb_table(mdb_file, table):
    export_res = subprocess.run(['mdb-export', mdb_file, table], stdout=subprocess.PIPE)
    io_str = io.StringIO(initial_value=export_res.stdout.decode('utf8'))
    items = read_csv(io_str)
    items_list = []
    for i, item in items.iterrows():
        items_list.append(dict(item))
    return items_list


def filter_by_page_nb(page_nb, items):
    return filter(
        lambda item: item["PageNo"] == page_nb,
        items
    )


def get_normed_image_path(image, replace_rules):
    path = image["ImagePath"]
    if not isinstance(path, str):
        raise IncoherentDataError("Image path must be a string, got '{}' instead.".format(type(path)))
    for rule in replace_rules:
        if path.startswith(rule[0]):
            path = path.replace(rule[0], "")
            while path.startswith('\\'): path = path[1:]
            path = path.replace('\\', '/')
            path = os.path.join(rule[1], path)
            return path
    else:
        raise ValueError("Could not find a replacement path for file '{}'.".format(path))


def draw_image(canvas, image):
    """Hypothèse: LeftPos, TopPos, Height, Width : position et taille de la zone d'image
    LastLeft et LastTop : distance entre le coin de l'image et le coin de la zone d'image
    LastHeight et lastWidth : Taille réelle de l'image
    """
    try:
        image_path = get_normed_image_path(image, PATH_REPLACE_RULES)
    except IncoherentDataError as e:
        print("WARNING: Could not load image '{book_image_id}' for page '{page_no}', at path '{image_path}'. "
              "Got IncoherentDataError : {msg}.".format(
            book_image_id=image["BookImageId"],
            page_no=image["PageNo"],
            image_path=image["ImagePath"],
            msg=str(e)
        ))
        return # skip image

    zone_top_left_corner_coords = (image["LeftPos"], PAGE_SIZE[1] - image["TopPos"])
    zone_bottom_left_corner_coords = (zone_top_left_corner_coords[0], zone_top_left_corner_coords[1] - image['Height'])

    image_top_left_corner_coords = (
    zone_top_left_corner_coords[0] - image['LastLeft'], zone_top_left_corner_coords[1] + image['LastTop'])
    image_bottom_left_corner_coords = (
    image_top_left_corner_coords[0], image_top_left_corner_coords[1] - image['LastHeight'])

    clipping_path = canvas.beginPath()
    clipping_path.rect(x=zone_bottom_left_corner_coords[0], y=zone_bottom_left_corner_coords[1],
                       width=image['Width'], height=image['Height'])

    canvas.saveState()
    if CLIP_IMAGES: canvas.clipPath(clipping_path, stroke=0)
    canvas.translate(image_bottom_left_corner_coords[0]+image['LastWidth']/2, image_bottom_left_corner_coords[1]+image['LastHeight']/2)

    rotation_x_margin = 0
    rotation_y_margin = 0
    effective_rotation_x_margin = 0
    effective_rotation_y_margin = 0
    if DO_ROTATE:
        rotation_angle = 360-image['ImageRotationAngle']
        canvas.rotate(rotation_angle)
        if ZOOM_ROTATE:
            if rotation_angle >= 0:
                rotation_x_margin = compute_zomming_margin_x(image, 360-image['ImageRotationAngle'])
                rotation_y_margin = (image['LastHeight']/image['LastWidth']) * rotation_x_margin
            else:
                rotation_y_margin = compute_zomming_margin_y(image, 360-image['ImageRotationAngle'])
                rotation_x_margin = (image['LastWidth']/image['LastHeight']) * rotation_y_margin

            natural_x_margin = (image['LastWidth'] - image['Width']) / 2
            natural_y_margin = (image['LastHeight'] - image['Height']) / 2

            if rotation_x_margin - natural_x_margin <= 0 and rotation_y_margin - natural_y_margin <= 0:
                pass # do nothing
            if rotation_x_margin - natural_x_margin <= 0 or rotation_y_margin - natural_y_margin <= 0:
                effective_rotation_x_margin = rotation_x_margin
                effective_rotation_y_margin = rotation_y_margin
            else:
                if rotation_x_margin - natural_x_margin > rotation_y_margin - natural_y_margin:
                    effective_rotation_x_margin = rotation_x_margin - natural_x_margin
                    effective_rotation_y_margin = (image['LastHeight']/image['LastWidth']) * effective_rotation_x_margin
                else:
                    effective_rotation_y_margin = rotation_y_margin - natural_y_margin
                    effective_rotation_x_margin = (image['LastWidth']/image['LastHeight']) * effective_rotation_y_margin

            print("(x, y): ({}, {})".format(rotation_y_margin, rotation_x_margin))

    canvas.drawImage(image_path, -image['LastWidth']/2 - effective_rotation_x_margin, -image['LastHeight']/2 - effective_rotation_y_margin,
                     width=image['LastWidth'] + 2*effective_rotation_x_margin, height=image['LastHeight'] + 2*effective_rotation_y_margin)
    canvas.restoreState()
    if DRAW_BORDER: canvas.drawPath(clipping_path, stroke=1)
    
    
def compute_zomming_margin_x(image, rotation_d):
    rotation_r = math.radians(rotation_d)
    top_right_coord = (image['Width']/2, image['Height']/2)
    diag = math.sqrt((image['Height'])**2 + (image['Width'])**2)

    # tan(a) = op/adj => a = arctan(op/adj)
    diag_angle = math.atan(image["Height"]/image["Width"])
    x_diff = (math.cos(diag_angle + rotation_r) - math.cos(diag_angle)) * diag
    y_diff = (math.sin(diag_angle + rotation_r) - math.sin(diag_angle)) * diag
    new_top_right_coord = (top_right_coord[0]+x_diff, top_right_coord[1]+y_diff)

    bottom_right_coord = (image['Width']/2, -image['Height']/2)
    new_bottom_right_coord = (bottom_right_coord[0]+y_diff, bottom_right_coord[1]-x_diff)
    a = norm(new_bottom_right_coord, new_top_right_coord)
    b = norm(new_bottom_right_coord, top_right_coord)
    c = norm(new_top_right_coord, top_right_coord)
    p = (a+b+c)/2
    area = math.sqrt(p*(p-a)*(p-b)*(p-c))

    # area = (1/2) * h * a => h = (area * 2) / a
    return (area * 2) / a

def compute_zomming_margin_y(image, rotation_d):
    rotation_r = math.radians(rotation_d)
    top_right_coord = (image['Width']/2, image['Height']/2)
    diag = math.sqrt((image['Height'])**2 + (image['Width'])**2)

    # tan(a) = op/adj => a = arctan(op/adj)
    diag_angle = math.atan(image["Height"]/image["Width"])
    x_diff = (math.cos(diag_angle + rotation_r) - math.cos(diag_angle)) * diag
    y_diff = (math.sin(diag_angle + rotation_r) - math.sin(diag_angle)) * diag
    new_top_right_coord = (top_right_coord[0]+x_diff, top_right_coord[1]+y_diff)

    top_left_corner = (image['Width']/2, -image['Height']/2)
    new_top_left_corner = (top_left_corner[0]+y_diff, top_left_corner[1]+x_diff)
    a = norm(new_top_left_corner, new_top_right_coord)
    b = norm(new_top_left_corner, top_right_coord)
    c = norm(new_top_right_coord, top_right_coord)
    p = (a+b+c)/2
    area = math.sqrt(p*(p-a)*(p-b)*(p-c))

    # area = (1/2) * h * a => h = (area * 2) / a
    return (area * 2) / a


def norm(p1, p2):
    return math.sqrt((p1[0]-p2[0])**2 + (p1[1]-p2[1])**2)


def draw_text(canvas, text):
    paragraph = parse_mapalb_xml_text(text['BText'])
    effective_font_size = paragraph["style"]["font_size"]*FONT_SIZE_CORRECT_FACTOR

    text_top_left_corner_coords = (text["LeftPos"], PAGE_SIZE[1] - text["TopPos"])
    text_bottom_left_corner_coords = (text_top_left_corner_coords[0], text_top_left_corner_coords[1] - text["Height"])

    # canvas.saveState()
    # canvas.setFillColorRGB(1, 0, 0)
    # canvas.rect(text_bottom_left_corner_coords[0], text_bottom_left_corner_coords[1]-(7/30)*parsed_text["style"]["font_size"],
    #             text["Width"], parsed_text["style"]["font_size"],
    #             fill=1, stroke=0)
    # canvas.restoreState()

    # rect showing the position where we are drawing
    # canvas.saveState()
    # canvas.setFillColorRGB(1, 0, 0)
    # canvas.rect(text_bottom_left_corner_coords[0]+LEFT_POS_CORRECT, text_bottom_left_corner_coords[1]+BOTTOM_POS_CORRECT,
    #             text["Width"], paragraph["style"]["font_size"],
    #             fill=1, stroke=0)
    # canvas.restoreState()

    # rect showing the box described in the data
    # canvas.saveState()
    # canvas.setFillColorRGB(1, 0, 0)
    # canvas.rect(text_bottom_left_corner_coords[0], text_bottom_left_corner_coords[1],
    #             text["Width"], text["Height"],
    #             fill=1, stroke=0)
    # canvas.restoreState()

    canvas.saveState()
    canvas.setFont("book-antiqua-bold", effective_font_size)
    canvas.setFillColorRGB(
        paragraph["style"]["colour"]["red"],
        paragraph["style"]["colour"]["green"],
        paragraph["style"]["colour"]["blue"]
    )

    text_to_draw_height = len(paragraph["lines"]) * paragraph["style"]["font_size"] # no v space between lines

    y = text_bottom_left_corner_coords[1] + (text['Height'] - text_to_draw_height) / 2 + effective_font_size * (7/30)
    for line in paragraph["lines"]:
        line_to_draw_width = canvas.stringWidth(line)
        x = text_bottom_left_corner_coords[0] + (text['Width'] - line_to_draw_width) / 2 # center text horizontally
        canvas.drawString(x=x, y=y, text=line)
        y += effective_font_size

    canvas.restoreState()

    # canvas.beginText()
    # canvas.drawCentredString()


def html_colour_to_rgb(html_colour):
    return {
        "red": int(html_colour[1:3], 16) / 255,
        "green": int(html_colour[3:5], 16) / 255,
        "blue": int(html_colour[5:7], 16) / 255,
    }


def parse_mapalb_xml_text(xml_str):
    attributes_to_copy = ["TextAlignment", "FontSize", "Foreground"]
    out = {
        "style": {
            "font_size": None,
            "color": None,
            "alignment": None
        },
        "lines": []
    }

    def update_style_from_attribs(style, xml_elt):
        if "FontSize" in xml_elt.attrib:
            style["font_size"] = int(xml_elt.attrib["FontSize"])
        if "Foreground" in xml_elt.attrib:
            style["colour"] = html_colour_to_rgb(xml_elt.attrib["Foreground"])
        if "TextAlignment" in xml_elt.attrib:
            style['alignment'] = xml_elt.attrib["TextAlignment"]
        return style

    root = ET.fromstring(xml_str)
    out["style"] = update_style_from_attribs(out["style"], root)
    for paragraph in root.iter('{http://schemas.microsoft.com/winfx/2006/xaml/presentation}Paragraph'):
        out["style"] = update_style_from_attribs(out["style"], paragraph)
        for run in paragraph.iter('{http://schemas.microsoft.com/winfx/2006/xaml/presentation}Run'):
            out["style"] = update_style_from_attribs(out["style"], run)
            out["lines"].append(run.text)

    if out["style"]["font_size"] is None:
        print("WARNING: no font_size defined for text '{}'.".format('\n'.join(out["lines"])))
        out["style"]["font_size"] = 14
    if out["style"]["colour"] is None:
        print("WARNING: no font_size defined for text '{}'.".format('\n'.join(out["lines"])))
        out["style"]["colour"] = {
            "red": 0,
            "green": 0,
            "blue": 0
        }

    return out


def get_page_bg_colour(page):
    return {
        "red": page["BackColor_Red"] / 255,
        "green": page["BackColor_Green"] / 255,
        "blue": page["BackColor_Blue"] / 255
    }


def register_fonts():
    pdfmetrics.registerFont(TTFont('book-antiqua', 'fonts/book-antiqua.ttf'))
    pdfmetrics.registerFont(TTFont('book-antiqua-bold', 'fonts/book-antiqua-bold.ttf'))


def run_script():
    """The script is defined inside a function so that all variables defined there are local"""
    args = parse_my_args()

    shutil.rmtree('temp')
    os.mkdir('temp')

    print("Loading data")
    pages = load_mdb_table(args.input_file, 'BookPages')
    images = load_mdb_table(args.input_file, 'BookImage')
    texts = load_mdb_table(args.input_file, 'BookText')
    print("Done.")

    n_pdf_file = 0
    cur_pdf_filename = args.output_file + str(n_pdf_file) + ".pdf"
    c = canvas.Canvas(cur_pdf_filename, pagesize=PAGE_SIZE)
    register_fonts()

    n_page_drawn = 0
    for page in pages:
        page_nb = page["PageNo"]
        # 0 : deuxieme de couv
        # -1 : premiere de couv
        # -2 : quatrieme de couv
        if page_nb <= 0:
            continue
        if page_nb > 10:
            continue
        print("Drawing page {}".format(page_nb))

        # set bg colour
        bg_colour = get_page_bg_colour(page)
        c.setFillColorRGB(bg_colour["red"], bg_colour["green"], bg_colour["blue"])
        c.rect(0, 0, PAGE_SIZE[0], PAGE_SIZE[1], fill=1, stroke=0)

        # draw images
        cur_images = filter_by_page_nb(page_nb, images)
        for image in cur_images:
            draw_image(c, image)

        # draw text
        cur_texts = filter_by_page_nb(page_nb, texts)
        for text in cur_texts:
            draw_text(c, text)

        c.showPage()
        n_page_drawn += 1
        print("Done.")

        if n_page_drawn % 10 == 0 and n_page_drawn != 0:
            print("Saving pdf file after page #{}. File : {}".format(n_page_drawn, cur_pdf_filename))
            c.save()
            print("Done.")

            n_pdf_file += 1
            cur_pdf_filename = args.output_file + str(n_pdf_file) + ".pdf"
            del c
            c = canvas.Canvas(cur_pdf_filename, pagesize=PAGE_SIZE)

    print("Saving file")
    c.save()
    print("Done.")


if __name__ == "__main__":
    run_script()
