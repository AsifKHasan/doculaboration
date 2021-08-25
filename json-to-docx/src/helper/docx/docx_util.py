#!/usr/bin/env python3

'''
various utilities for formatting a docx
'''

import lxml

from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor, Emu
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.table import _Cell

from copy import deepcopy

from helper.logger import *

GSHEET_OXML_BORDER_MAPPING = {
    'DOTTED': 'dotted',
    'DASHED': 'dashed',
    'SOLID': 'single',
    'SOLID_MEDIUM': 'thick',
    'SOLID_THICK': 'triple',
    'DOUBLE': 'double',
    'NONE': 'none'
}

def set_updatefields_true(docx_path):
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    doc = Document(docx_path)
    # add child to doc.settings element
    element_updatefields = lxml.etree.SubElement(
        doc.settings.element, f"{namespace}updateFields"
    )
    element_updatefields.set(f"{namespace}val", "true")
    doc.save(docx_path)

'''
Table of Contents
'''
def add_toc(doc):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    # creates a new element
    fldChar = OxmlElement('w:fldChar')
    # sets attribute on element
    fldChar.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    # sets attribute on element
    instrText.set(qn('xml:space'), 'preserve')
    # change 1-3 depending on heading levels you need
    instrText.text = 'TOC \\o "1-6" \\h \\z \\u'

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:t')
    fldChar3.text = "Right-click to update Table of Contents."
    fldChar2.append(fldChar3)

    fldChar4 = OxmlElement('w:fldChar')
    fldChar4.set(qn('w:fldCharType'), 'end')

    r_element = run._r
    r_element.append(fldChar)
    r_element.append(instrText)
    r_element.append(fldChar2)
    r_element.append(fldChar4)
    p_element = paragraph._p

'''
List of Figures
'''
def add_lof(doc):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    # creates a new element
    fldChar = OxmlElement('w:fldChar')
    # sets attribute on element
    fldChar.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    # sets attribute on element
    instrText.set(qn('xml:space'), 'preserve')
    # change 1-3 depending on heading levels you need
    instrText.text = 'TOC \\h \\z \\t "Figure" \\c'

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:t')
    fldChar3.text = "Right-click to update List of Figures."
    fldChar2.append(fldChar3)

    fldChar4 = OxmlElement('w:fldChar')
    fldChar4.set(qn('w:fldCharType'), 'end')

    r_element = run._r
    r_element.append(fldChar)
    r_element.append(instrText)
    r_element.append(fldChar2)
    r_element.append(fldChar4)
    p_element = paragraph._p

'''
List of Tables
'''
def add_lot(doc):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    # creates a new element
    fldChar = OxmlElement('w:fldChar')
    # sets attribute on element
    fldChar.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    # sets attribute on element
    instrText.set(qn('xml:space'), 'preserve')
    # change 1-3 depending on heading levels you need
    instrText.text = 'TOC \\h \\z \\t "Table" \\c'

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:t')
    fldChar3.text = "Right-click to update List of Figures."
    fldChar2.append(fldChar3)

    fldChar4 = OxmlElement('w:fldChar')
    fldChar4.set(qn('w:fldCharType'), 'end')

    r_element = run._r
    r_element.append(fldChar)
    r_element.append(instrText)
    r_element.append(fldChar2)
    r_element.append(fldChar4)
    p_element = paragraph._p

def add_numbered_paragraph(doc, text, restart=False, style='List Number'):
    if restart:
        # new numbering
        ABSTRACT_NUM_ID = 8
        numbering = doc._part.numbering_part.numbering_definitions._numbering
        num_id = numbering._next_numId
        num = CT_Num.new(num_id, ABSTRACT_NUM_ID)
        num.add_lvlOverride(ilvl=0).add_startOverride(1)
        w_num = numbering._insert_num(num)

        #w_num_pr_xml = '<w:ilvl w:val="0"/><w:numId w:val="{0}"/>'.format(num_id)
        #w_num_pr = CT_NumPr()
        p = doc.add_paragraph(text, style)

         # Access paragraph XML element
        p_xml = p._p

        # Paragraph properties
        p_props = p_xml.get_or_add_pPr()

        # Create number properties element
        num_props = OxmlElement('w:numPr')

        lvl_prop = OxmlElement('w:ilvl')
        lvl_prop.set(qn('w:val'), '0')

        num_id_prop = OxmlElement('w:numId')
        num_id_prop.set(qn('w:val'), str(w_num.numId))

        num_props.append(lvl_prop)
        num_props.append(num_id_prop)

        # Add number properties to paragraph
        p_props.append(num_props)

    else:
        doc.add_paragraph(text, style)

def add_horizontal_line(paragraph, pos='w:bottom', size='6', color='auto'):
    p = paragraph._p  # p is the <w:p> XML element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert_element_before(pBdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
        'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
        'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
        'w:pPrChange'
    )
    bottom = OxmlElement(pos)
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), size)
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)

def append_page_number_only(paragraph):
    run = paragraph.add_run()

    # page number
    # creates a new element
    fldCharBegin1 = OxmlElement('w:fldChar')
    # sets attribute on element
    fldCharBegin1.set(qn('w:fldCharType'), 'begin')
    instrText1 = OxmlElement('w:instrText')
    # sets attribute on element
    instrText1.set(qn('xml:space'), 'preserve')
    # page number
    instrText1.text = 'PAGE \* MERGEFORMAT'

    fldCharSeparate1 = OxmlElement('w:fldChar')
    fldCharSeparate1.set(qn('w:fldCharType'), 'separate')

    fldCharEnd1 = OxmlElement('w:fldChar')
    fldCharEnd1.set(qn('w:fldCharType'), 'end')

    r_element = run._r
    r_element.append(fldCharBegin1)
    r_element.append(instrText1)
    r_element.append(fldCharSeparate1)
    r_element.append(fldCharEnd1)
    p_element = paragraph._p

def append_page_number_with_pages(paragraph, separator=' of '):
    run = paragraph.add_run()

    # page number
    # creates a new element
    fldCharBegin1 = OxmlElement('w:fldChar')
    # sets attribute on element
    fldCharBegin1.set(qn('w:fldCharType'), 'begin')
    instrText1 = OxmlElement('w:instrText')
    # sets attribute on element
    instrText1.set(qn('xml:space'), 'preserve')
    # page number
    instrText1.text = 'PAGE \* MERGEFORMAT'

    fldCharSeparate1 = OxmlElement('w:fldChar')
    fldCharSeparate1.set(qn('w:fldCharType'), 'separate')

    fldCharEnd1 = OxmlElement('w:fldChar')
    fldCharEnd1.set(qn('w:fldCharType'), 'end')

    fldCharOf = OxmlElement('w:t')
    fldCharOf.set(qn('xml:space'), 'preserve')
    fldCharOf.text = separator

    # number of pages
    # creates a new element
    fldCharBegin2 = OxmlElement('w:fldChar')
    # sets attribute on element
    fldCharBegin2.set(qn('w:fldCharType'), 'begin')
    instrText2 = OxmlElement('w:instrText')
    # sets attribute on element
    instrText2.set(qn('xml:space'), 'preserve')
    # page number
    instrText2.text = 'NUMPAGES \* MERGEFORMAT'

    fldCharSeparate2 = OxmlElement('w:fldChar')
    fldCharSeparate2.set(qn('w:fldCharType'), 'separate')

    fldCharEnd2 = OxmlElement('w:fldChar')
    fldCharEnd2.set(qn('w:fldCharType'), 'end')

    r_element = run._r
    r_element.append(fldCharBegin1)
    r_element.append(instrText1)
    r_element.append(fldCharSeparate1)
    r_element.append(fldCharEnd1)
    r_element.append(fldCharOf)
    r_element.append(fldCharBegin2)
    r_element.append(instrText2)
    r_element.append(fldCharSeparate2)
    r_element.append(fldCharEnd2)
    p_element = paragraph._p

def rotate_text(cell: _Cell, direction: str):
    # direction: tbRl -- top to bottom, btLr -- bottom to top
    assert direction in ("tbRl", "btLr")
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    textDirection = OxmlElement('w:textDirection')
    textDirection.set(qn('w:val'), direction)  # btLr tbRl
    tcPr.append(textDirection)

def set_character_style(run, spec):
    run.bold = spec['bold']
    run.italic = spec['italic']
    run.strike = spec['strikethrough']
    run.underline = spec['underline']

    run.font.name = spec['fontFamily']
    run.font.size = Pt(spec['fontSize'])

    red, green, blue = 0, 0, 0
    fgcolor = spec['foregroundColor']
    if fgcolor == {}:
        red, green, blue = 32, 32, 32
    else:
        red = int(fgcolor['red'] * 255) if 'red' in fgcolor else 0
        green = int(fgcolor['green'] * 255) if 'green' in fgcolor else 0
        blue = int(fgcolor['blue'] * 255) if 'blue' in fgcolor else 0

    run.font.color.rgb = RGBColor(red, green, blue)

def set_cell_bgcolor(cell, color):
    shading_elm_1 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color))
    cell._tc.get_or_add_tcPr().append(shading_elm_1)

def set_paragraph_bgcolor(paragraph, color):
    shading_elm_1 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color))
    paragraph._p.get_or_add_pPr().append(shading_elm_1)

def copy_cell_border(from_cell: _Cell, to_cell: _Cell):
    from_tc = from_cell._tc
    from_tcPr = from_tc.get_or_add_tcPr()

    to_tc = to_cell._tc
    to_tcPr = to_tc.get_or_add_tcPr()

    from_tcBorders = from_tcPr.first_child_found_in("w:tcBorders")
    to_tcBorders = to_tcPr.first_child_found_in("w:tcBorders")
    if from_tcBorders is not None:
        if to_tcBorders is None:
            to_tcBorders = deepcopy(from_tcBorders)
            to_tc.get_or_add_tcPr().append(to_tcBorders)

def set_cell_border(cell: _Cell, **kwargs):
    """
    Set cell's border
    Usage:

    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def set_paragraph_border(paragraph, **kwargs):
    """
    Set paragraph's border
    Usage:

    set_paragraph_border(
        paragraph,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )
    """
    pPr = paragraph._p.get_or_add_pPr()

    # check for tag existnace, if none found, then create one
    pBorders = pPr.first_child_found_in("w:pBorders")
    if pBorders is None:
        pBorders = OxmlElement('w:pBorders')
        pPr.append(pBorders)

    # list over all available tags
    for edge in ('top', 'start', 'bottom', 'end'):
        edge_data = kwargs.get(edge)
        if edge_data:
            edge_str = edge
            if edge_str == 'start': edge_str = 'left'
            if edge_str == 'end': edge_str = 'right'
            tag = 'w:{}'.format(edge_str)

            # check for tag existnace, if none found, then create one
            element = pBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                pBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def merge_document(placeholder, docx_path):
    # the document is in the same folder as the template, get the path
    # docx_path = os.path.abspath('{0}/{1}'.format('../conf', docx_name))
    sub_doc = Document(docx_path)

    par_parent = placeholder._p.getparent()
    index = par_parent.index(placeholder._p) + 1
    for element in sub_doc.part.element:
        element.remove_all('w:sectPr')
        par_parent.insert(index, element)
        index = index + 1

def polish_table(table):
    for r in table.rows:
        # no preferred width for the last column
        c = r._tr.tc_lst[-1]
        #for c in r._tr.tc_lst:
        tcW = c.tcPr.tcW
        tcW.type = 'auto'
        tcW.w = 0

def ooxml_border_from_gsheet_border(borders, key):
    if key in borders:
        border = borders[key]
        red = int(border['color']['red'] * 255) if 'red' in border['color'] else 0
        green = int(border['color']['green'] * 255) if 'green' in border['color'] else 0
        blue = int(border['color']['blue'] * 255) if 'blue' in border['color'] else 0
        color = '{0:02x}{1:02x}{2:02x}'.format(red, green, blue)
        if 'style' in border:
            border_style = border['style']
        else:
            border_style = 'NONE'

        return {"sz": border['width'] * 8, "val": GSHEET_OXML_BORDER_MAPPING[border_style], "color": color, "space": "0"}
    else:
        return None

def insert_image(cell, image_spec):
    '''
        image_spec is like {'url': url, 'path': local_path, 'height': height, 'width': width, 'dpi': im_dpi}
    '''
    if image_spec is not None:
        run = cell.paragraphs[0].add_run()
        run.add_picture(image_spec['path'], height=Pt(image_spec['height']), width=Pt(image_spec['width']))

def set_repeat_table_header(row):
    ''' set repeat table row on every new page
    '''
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    tblHeader.set(qn('w:val'), "true")
    trPr.append(tblHeader)
    return row
