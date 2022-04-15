#!/usr/bin/env python3

''' various utilities for generating an Openoffice odt document
'''
import platform
import subprocess
import random
import string
from pathlib import Path

from odf.style import Style, MasterPage, PageLayout, PageLayoutProperties, TextProperties, ParagraphProperties, TableProperties, TableColumnProperties, TableRowProperties, TableCellProperties
from odf.text import P, Span, TableOfContent, TableOfContentSource, IndexTitleTemplate
from odf.table import Table, TableColumns, TableColumn, TableRows, TableRow, TableCell

from helper.logger import *


if platform.system() == 'Windows':
    LIBREOFFICE_EXECUTABLE = 'C:/Program Files/LibreOffice/program/soffice.exe'
else:
    LIBREOFFICE_EXECUTABLE = 'soffice'


# --------------------------------------------------------------------------------------------------------------------------------------------
# table, table-row, table-column, table-cell

''' create a Table
'''
def create_table(odt, table_style_attributes, table_properties_attributes):
    if 'family' not in table_style_attributes:
        table_style_attributes['family'] = 'table'

    # create the style
    table_style = Style(attributes=table_style_attributes)
    table_style.addElement(TableProperties(attributes=table_properties_attributes))
    odt.automaticstyles.addElement(table_style)

    # create the table
    table_properties = {'name': table_style.getAttribute('name'), 'stylename': table_style.getAttribute('name')}
    table = Table(attributes=table_properties)

    return table


''' create TableColumn
'''
def create_table_column(odt, table_column_style_attributes, table_column_properties_attributes):
    if 'family' not in table_column_style_attributes:
        table_column_style_attributes['family'] = 'table-column'

    # create the style
    table_column_style = Style(attributes=table_column_style_attributes)
    table_column_style.addElement(TableColumnProperties(attributes=table_column_properties_attributes))
    odt.automaticstyles.addElement(table_column_style)

    # create the table-column
    table_column_properties = {'stylename': table_column_style.getAttribute('name')}
    table_column = TableColumn(attributes=table_column_properties)

    return table_column


''' create TableRow
'''
def create_table_row(odt, table_row_style_attributes, table_row_properties_attributes):
    if 'family' not in table_row_style_attributes:
        table_row_style_attributes['family'] = 'table-row'

    # create the style
    table_row_style = Style(attributes=table_row_style_attributes)
    table_row_style.addElement(TableRowProperties(attributes=table_row_properties_attributes))
    odt.automaticstyles.addElement(table_row_style)

    # create the table-row
    table_row_properties = {'stylename': table_row_style.getAttribute('name')}
    table_row = TableRow(attributes=table_row_properties)

    return table_row


''' create TableCell
'''
def create_table_cell(odt, table_cell_style_attributes, table_cell_properties_attributes, table_cell_attributes):
    if 'family' not in table_cell_style_attributes:
        table_cell_style_attributes['family'] = 'table-cell'

    # create the style
    table_cell_style = Style(attributes=table_cell_style_attributes)
    table_cell_style.addElement(TableCellProperties(attributes=table_cell_properties_attributes))
    odt.automaticstyles.addElement(table_cell_style)

    # create the table-cell
    table_cell_properties = {'stylename': table_cell_style.getAttribute('name')} | table_cell_attributes
    table_cell = TableCell(attributes=table_cell_properties)

    return table_cell


# --------------------------------------------------------------------------------------------------------------------------------------------
# paragraphs and styles

''' create style - family paragraph
'''
def create_paragraph_style(odt, style_attributes=None, paragraph_attributes=None, text_attributes=None):
    # we may need to create the style-name
    if style_attributes == None:
        style_attributes = {}

    if 'family' not in style_attributes:
        style_attributes['family'] = 'paragraph'

    if 'name' not in style_attributes:
        style_attributes['name'] = random_string()

    if 'parentstylename' not in style_attributes:
        style_attributes['parentstylename'] = 'Text_20_body'

    # create the style
    paragraph_style = Style(attributes=style_attributes)

    if paragraph_attributes is not None:
        paragraph_style.addElement(ParagraphProperties(attributes=paragraph_attributes))

    if text_attributes is not None:
        paragraph_style.addElement(TextProperties(attributes=text_attributes))

    odt.automaticstyles.addElement(paragraph_style)

    return paragraph_style.getAttribute('name')



''' write a paragraph in a given style
'''
def create_paragraph(odt, style_name, text=None, run_list=None):
    style = odt.getStyleByName(style_name)
    if text is not None:
        paragraph = P(stylename=style, text=text)
        return paragraph

    if run_list is not None:
        paragraph = P(stylename=style)
        for run in run_list:
            style_attributes = {'family': 'text'}
            text_style_name = create_paragraph_style(odt, style_attributes=style_attributes, text_attributes=run['text-attributes'])
            paragraph.addElement(Span(stylename=text_style_name, text=run['text']))

        return paragraph



# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# indexes and pdf generation

''' update indexes through a macro macro:///Standard.Module1.open_document(document_url) which must be in OpenOffice macro library
'''
def update_indexes(odt, odt_path):
    document_url = Path(odt_path).as_uri()
    macro = f'"macro:///Standard.Module1.open_document("{document_url}")"'
    command_line = f'"{LIBREOFFICE_EXECUTABLE}" --headless --invisible {macro}'
    debug(f"updating indexes for {odt_path}")
    subprocess.call(command_line, shell=True);


''' given an odt file generates pdf in the given directory
'''
def generate_pdf(infile, outdir):
    command_line = f'"{LIBREOFFICE_EXECUTABLE}" --headless --convert-to pdf --outdir "{outdir}" "{infile}"'
    debug(f"generating pdf from {infile}")
    subprocess.call(command_line, shell=True);


''' create table-of-contents
'''
def create_toc(odt):
    name = 'Table of Content'
    toc = TableOfContent(name=name)
    toc_source = TableOfContentSource(outlinelevel=10)
    toc_title_template = IndexTitleTemplate()
    toc_source.addElement(toc_title_template)
    toc.addElement(toc_source)
    odt.text.addElement(toc)


''' create illustration-index
'''
def create_lof(odt):
    name = 'List of Figures'


''' create Table-index
'''
def create_lot(odt):
    name = 'List of Tables'



# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# master-page and page-layout

''' update page-layout of Standard master-page with the given page-layout
'''
def update_standard_master_page(odt, new_page_layout):
    # get the Standard master-page
    standard_master_page_name = 'Standard'
    standard_master_page = None
    for master_page in odt.masterstyles.getElementsByType(MasterPage):
        if master_page.getAttribute('name') == standard_master_page_name:
            standard_master_page = master_page
            break

    if standard_master_page is not None:
        debug(f"master-page {standard_master_page.getAttribute('name')} found; page-layout is {standard_master_page.attributes[(standard_master_page.qname[0], 'page-layout-name')]}")
        standard_master_page.attributes[(standard_master_page.qname[0], 'page-layout-name')] = new_page_layout
        debug(f"master-page {standard_master_page.getAttribute('name')}  page-layout changed to {standard_master_page.attributes[(standard_master_page.qname[0], 'page-layout-name')]}")
    else:
        warn(f"master-page {standard_master_page_name} NOT found")


''' gets the page-layout from odt if it is there, else create one
'''
def get_or_create_page_layout(odt, odt_specs, page_layout_name, page_spec, margin_spec, orientation):
    page_layout = None

    for pl in odt.automaticstyles.getElementsByType(PageLayout):
        if pl.getAttribute('name') == page_layout_name:
            page_layout = pl
            break

    if page_layout is None:
        # create one
        page_layout = PageLayout(name=page_layout_name)
        odt.automaticstyles.addElement(page_layout)
        pageheight = f"{odt_specs['page-spec'][page_spec]['height']}in"
        pagewidth = f"{odt_specs['page-spec'][page_spec]['width']}in"
        margintop = f"{odt_specs['margin-spec'][margin_spec]['top']}in"
        marginbottom = f"{odt_specs['margin-spec'][margin_spec]['bottom']}in"
        marginleft = f"{odt_specs['margin-spec'][margin_spec]['left']}in"
        marginright = f"{odt_specs['margin-spec'][margin_spec]['right']}in"
        # margingutter = f"{odt_specs['margin-spec'][margin_spec]['gutter']}in"
        printorientation = orientation
        page_layout.addElement(PageLayoutProperties(pagewidth=pagewidth, pageheight=pageheight, margintop=margintop, marginbottom=marginbottom, marginleft=marginleft, marginright=marginright, printorientation=orientation))

    return page_layout


''' new odt master-page
    page layouts are saved with a name page-spec__margin-spec__[portrait|landscape]
'''
def get_or_create_master_page(odt, odt_specs, page_layout_name, page_spec, margin_spec, orientation):
    # see if the master-page already exists or not in the odt
    master_page = None

    for mp in odt.masterstyles.childNodes:
        if mp.getAttribute('name') == page_layout_name:
            master_page = mp
            break

    if master_page is None:
        # create one, first get/create the page-layout
        page_layout = get_or_create_page_layout(odt, odt_specs, page_layout_name, page_spec, margin_spec, orientation)
        master_page = MasterPage(name=page_layout_name, pagelayoutname=page_layout_name)
        odt.masterstyles.addElement(master_page)

    return master_page


COLSEP = (6/72)
ROWSEP = (2/72)

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# various utility functions

''' given pixel size, calculate the row height in inches
    a reasonable approximation is what gsheet says 21 pixels, renders well as 12 pixel (assuming our normal text is 10-11 in size)
'''
def row_height_in_inches(pixel_size):
    return (pixel_size - 9) / 72


''' get a random string
'''
def random_string(length=6):
    letters = string.ascii_uppercase
    return ''.join(random.choice(letters) for i in range(length))



# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# various utility data

GSHEET_LATEX_BORDER_MAPPING = {
    'DOTTED': 'dotted',
    'DASHED': 'dashed',
    'SOLID': 'solid'
}


TEXT_VALIGN_MAP = {'TOP': 'top', 'MIDDLE': 'middle', 'BOTTOM': 'bottom'}
TEXT_HALIGN_MAP = {'LEFT': 'start', 'CENTER': 'center', 'RIGHT': 'end', 'JUSTIFY': 'justify'}
# TEXT_HALIGN_MAP = {'LEFT': 'left', 'CENTER': 'center', 'RIGHT': 'right', 'JUSTIFY': 'justify'}


COLSEP = (6/72)
ROWSEP = (2/72)

HEADER_FOOTER_FIRST_COL_HSPACE = -6
HEADER_FOOTER_LAST_COL_HSPACE = -6

COLUMNS = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
            'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
            'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ']
