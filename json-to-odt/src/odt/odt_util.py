#!/usr/bin/env python3

''' various utilities for generating an Openoffice odt document
'''
import platform
import subprocess
import random
import string
from pathlib import Path

from odf.style import Style, MasterPage, PageLayout, PageLayoutProperties, HeaderStyle, FooterStyle, HeaderFooterProperties, Header, HeaderLeft, Footer, FooterLeft
from odf.style import TextProperties, ParagraphProperties, TableProperties, TableColumnProperties, TableRowProperties, TableCellProperties, BackgroundImage
from odf.text import P, Span, TableOfContent, TableOfContentSource, IndexTitleTemplate, TableOfContentEntryTemplate, IndexSourceStyles, IndexSourceStyle, IndexEntryLinkStart, IndexEntryChapter, IndexEntryText, IndexEntryTabStop, IndexEntryPageNumber, IndexEntryLinkEnd
from odf.table import Table, TableColumns, TableColumn, TableRows, TableRow, TableCell

from helper.logger import *


if platform.system() == 'Windows':
    LIBREOFFICE_EXECUTABLE = 'C:/Program Files/LibreOffice/program/soffice.exe'
else:
    LIBREOFFICE_EXECUTABLE = 'soffice'


# --------------------------------------------------------------------------------------------------------------------------------------------
# pictures, background image

''' add a picture to the document
'''
def add_picture(odt, picture_path):
    href = odt.addPicture(picture_path)
    return href


''' create background-image-style
    <style:background-image xlink:href="Pictures/1000000100000115000000DD077AD77C6D58F87D.png" xlink:type="simple" xlink:actuate="onLoad" style:position="center center" style:repeat="no-repeat"/>
'''
def create_background_image_style(href, position):
    background_image_style = BackgroundImage(href=href, position=position, repeat='no-repeat', type='simple', actuate='onLoad')
    return background_image_style



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

    table_column_properties_attributes['useoptimalcolumnwidth'] = False

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
    # print(table_row_style_attributes)
    if 'family' not in table_row_style_attributes:
        table_row_style_attributes['family'] = 'table-row'

    # create the style
    table_row_style = Style(attributes=table_row_style_attributes)
    table_row_style.addElement(TableRowProperties(attributes=table_row_properties_attributes))
    odt.automaticstyles.addElement(table_row_style)

    # create the table-row
    table_row_properties = {'stylename': table_row_style.getAttribute('name')}
    # print(table_row_properties)
    table_row = TableRow(attributes=table_row_properties)

    return table_row


''' create TableCell
'''
def create_table_cell(odt, table_cell_style_attributes, table_cell_properties_attributes, table_cell_attributes, background_image_style=None):
    if 'family' not in table_cell_style_attributes:
        table_cell_style_attributes['family'] = 'table-cell'

    # create the style
    table_cell_style = Style(attributes=table_cell_style_attributes)
    table_cell_properties = TableCellProperties(attributes=table_cell_properties_attributes)

    if background_image_style:
        table_cell_properties.addElement(background_image_style)

    table_cell_style.addElement(table_cell_properties)
    odt.automaticstyles.addElement(table_cell_style)

    # create the table-cell
    table_cell_attributes['stylename'] = table_cell_style.getAttribute('name')
    table_cell = TableCell(attributes=table_cell_attributes)

    return table_cell


''' create CoveredTableCell
'''
def create_covered_table_cell(odt, table_cell_style_attributes, table_cell_properties_attributes):
    if 'family' not in table_cell_style_attributes:
        table_cell_style_attributes['family'] = 'table-cell'

    # create the style
    table_cell_style = Style(attributes=table_cell_style_attributes)
    table_cell_style.addElement(TableCellProperties(attributes=table_cell_properties_attributes))
    odt.automaticstyles.addElement(table_cell_style)

    # create the table-cell
    table_cell_properties = {'stylename': table_cell_style.getAttribute('name')}
    table_cell = TableCell(attributes=table_cell_properties)

    return table_cell



# --------------------------------------------------------------------------------------------------------------------------------------------
# paragraphs and styles

''' create style - family paragraph
'''
def create_paragraph_style(odt, style_attributes=None, paragraph_attributes=None, text_attributes=None):
    # we may need to create the style-name
    if style_attributes is None:
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
    if style is None:
        warn(f"style {style_name} not found")

    if text is not None:
        paragraph = P(stylename=style_name, text=text)
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
    debug(f"updating indexes for {Path(odt_path).resolve()}")
    subprocess.call(command_line, shell=True);


''' given an odt file generates pdf in the given directory
'''
def generate_pdf(infile, outdir):
    command_line = f'"{LIBREOFFICE_EXECUTABLE}" --headless --convert-to pdf --outdir "{outdir}" "{infile}"'
    debug(f"generating pdf from {Path(infile).resolve()}")
    subprocess.call(command_line, shell=True);


''' create table-of-contents
'''
def create_toc():
    name = 'Table of Content'
    toc = TableOfContent(name=name)
    toc_source = TableOfContentSource(outlinelevel=10)
    toc_title_template = IndexTitleTemplate()
    toc_source.addElement(toc_title_template)
    toc.addElement(toc_source)

    return toc


''' create illustration-index
'''
def create_lof():
    name = 'List of Figures'
    toc = TableOfContent(name=name)
    toc_source = TableOfContentSource(outlinelevel=1, useoutlinelevel=False, useindexmarks=False, useindexsourcestyles=True)
    toc_entry_template = TableOfContentEntryTemplate(outlinelevel=1, stylename='Figure')

    index_entry_link_start = IndexEntryLinkStart(stylename='Index_20_Link')
    index_entry_chapter = IndexEntryChapter()
    index_entry_text = IndexEntryText()
    index_entry_tab_stop = IndexEntryTabStop(type='right', leaderchar='.')
    index_entry_page_number = IndexEntryPageNumber()
    index_entry_link_end = IndexEntryLinkEnd()

    toc_entry_template.addElement(index_entry_link_start)
    toc_entry_template.addElement(index_entry_chapter)
    toc_entry_template.addElement(index_entry_text)
    toc_entry_template.addElement(index_entry_tab_stop)
    toc_entry_template.addElement(index_entry_page_number)
    toc_entry_template.addElement(index_entry_link_end)

    toc_title_template = IndexTitleTemplate()
    toc_index_source_styles = IndexSourceStyles(outlinelevel=1)
    toc_index_source_style = IndexSourceStyle(stylename='Figure')
    toc_index_source_styles.addElement(toc_index_source_style)

    toc_source.addElement(toc_entry_template)
    toc_source.addElement(toc_title_template)
    toc_source.addElement(toc_index_source_styles)
    toc.addElement(toc_source)

    return toc


''' create Table-index
'''
def create_lot():
    name = 'List of Tables'
    toc = TableOfContent(name=name)
    toc_source = TableOfContentSource(outlinelevel=1, useoutlinelevel=False, useindexmarks=False, useindexsourcestyles=True)
    toc_entry_template = TableOfContentEntryTemplate(outlinelevel=1, stylename='Table')

    index_entry_link_start = IndexEntryLinkStart(stylename='Index_20_Link')
    index_entry_chapter = IndexEntryChapter()
    index_entry_text = IndexEntryText()
    index_entry_tab_stop = IndexEntryTabStop(type='right', leaderchar='.')
    index_entry_page_number = IndexEntryPageNumber()
    index_entry_link_end = IndexEntryLinkEnd()

    toc_entry_template.addElement(index_entry_link_start)
    toc_entry_template.addElement(index_entry_chapter)
    toc_entry_template.addElement(index_entry_text)
    toc_entry_template.addElement(index_entry_tab_stop)
    toc_entry_template.addElement(index_entry_page_number)
    toc_entry_template.addElement(index_entry_link_end)

    toc_title_template = IndexTitleTemplate()
    toc_index_source_styles = IndexSourceStyles(outlinelevel=1)
    toc_index_source_style = IndexSourceStyle(stylename='Table')
    toc_index_source_styles.addElement(toc_index_source_style)

    toc_source.addElement(toc_entry_template)
    toc_source.addElement(toc_title_template)
    toc_source.addElement(toc_index_source_styles)
    toc.addElement(toc_source)

    return toc



# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# master-page and page-layout

''' get master-page by name
'''
def get_master_page(odt, master_page_name):
    for master_page in odt.masterstyles.getElementsByType(MasterPage):
        if master_page.getAttribute('name') == master_page_name:
            return master_page

    warn(f"master-page {master_page_name} NOT found")
    return None


''' get page-layout by name
'''
def get_page_layout(odt, page_layout_name):
    for page_layout in odt.automaticstyles.getElementsByType(PageLayout):
        if page_layout.getAttribute('name') == page_layout_name:
            return page_layout

    warn(f"page-layout {page_layout_name} NOT found")
    return None


''' update page-layout of Standard master-page with the given page-layout
'''
def update_master_page_page_layout(odt, master_page_name, new_page_layout_name):
    master_page = get_master_page(odt, master_page_name)

    if master_page is not None:
        master_page.attributes[(master_page.qname[0], 'page-layout-name')] = new_page_layout_name


''' create (section-specific) page-layout
'''
def create_page_layout(odt, odt_specs, page_layout_name, page_spec, margin_spec, orientation):
    # create one
    page_layout = PageLayout(name=page_layout_name)
    odt.automaticstyles.addElement(page_layout)
    printorientation = orientation
    if orientation == 'portrait':
        pageheight = f"{odt_specs['page-spec'][page_spec]['height']}in"
        pagewidth = f"{odt_specs['page-spec'][page_spec]['width']}in"
    else:
        pageheight = f"{odt_specs['page-spec'][page_spec]['width']}in"
        pagewidth = f"{odt_specs['page-spec'][page_spec]['height']}in"

    margintop = f"{odt_specs['margin-spec'][margin_spec]['top']}in"
    marginbottom = f"{odt_specs['margin-spec'][margin_spec]['bottom']}in"
    marginleft = f"{odt_specs['margin-spec'][margin_spec]['left']}in"
    marginright = f"{odt_specs['margin-spec'][margin_spec]['right']}in"
    # margingutter = f"{odt_specs['margin-spec'][margin_spec]['gutter']}in"
    page_layout.addElement(PageLayoutProperties(pagewidth=pagewidth, pageheight=pageheight, margintop=margintop, marginbottom=marginbottom, marginleft=marginleft, marginright=marginright, printorientation=orientation))

    return page_layout


''' create (section-specific) master-page
    page layouts are saved with a name mp-section-no
'''
def create_master_page(odt, odt_specs, master_page_name, page_layout_name, page_spec, margin_spec, orientation):
    # create one, first get/create the page-layout
    page_layout = create_page_layout(odt, odt_specs, page_layout_name, page_spec, margin_spec, orientation)
    master_page = MasterPage(name=master_page_name, pagelayoutname=page_layout_name)
    odt.masterstyles.addElement(master_page)

    return master_page


''' create header/footer
    <style:header-style>
        <style:header-footer-properties fo:min-height="0.0402in" fo:margin-left="0in" fo:margin-right="0in" fo:margin-bottom="0in" fo:background-color="transparent" style:dynamic-spacing="false" draw:fill="none"/>
    </style:header-style>
    <style:footer-style>
        <style:header-footer-properties svg:height="0.0402in" fo:margin-left="0in" fo:margin-right="0in" fo:margin-top="0in" fo:background-color="transparent" style:dynamic-spacing="false" draw:fill="none"/>
    </style:footer-style>
'''
def create_header_footer(master_page, page_layout, header_footer, odd_even):
    header_footer_style = None
    if header_footer == 'header':
        if odd_even == 'odd':
            header_footer_style = Header()
        if odd_even == 'even':
            header_footer_style = HeaderLeft()
        if odd_even == 'first':
            # header_footer_style = HeaderFirst()
            pass

    elif header_footer == 'footer':
        if odd_even == 'odd':
            header_footer_style = Footer()
        if odd_even == 'even':
            header_footer_style = FooterLeft()
        if odd_even == 'first':
            # header_footer_style = FooterFirst()
            pass

    if header_footer_style:
        header_footer_properties_attributes = {'margin': '0in', 'padding': '0in', 'dynamicspacing': False}
        header_style = HeaderStyle()
        header_style.addElement(HeaderFooterProperties(attributes=header_footer_properties_attributes))
        footer_style = FooterStyle()
        footer_style.addElement(HeaderFooterProperties(attributes=header_footer_properties_attributes))
        page_layout.addElement(header_style)
        page_layout.addElement(footer_style)

        master_page.addElement(header_footer_style)

    return header_footer_style


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# various utility functions

''' given pixel size, calculate the row height in inches
    a reasonable approximation is what gsheet says 21 pixels, renders well as 12 pixel (assuming our normal text is 10-11 in size)
'''
def row_height_in_inches(pixel_size):
    return (pixel_size) / 72


''' get a random string
'''
def random_string(length=12):
    letters = string.ascii_uppercase
    return ''.join(random.choice(letters) for i in range(length))



# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# various utility data

COLSEP = (0/72)
# ROWSEP = (2/72)

GSHEET_ODT_BORDER_MAPPING = {
    'DOTTED': 'dotted',
    'DASHED': 'dash',
    'SOLID': 'solid'
}


TEXT_VALIGN_MAP = {'TOP': 'top', 'MIDDLE': 'middle', 'BOTTOM': 'bottom'}
# TEXT_HALIGN_MAP = {'LEFT': 'start', 'CENTER': 'center', 'RIGHT': 'end', 'JUSTIFY': 'justify'}
TEXT_HALIGN_MAP = {'LEFT': 'left', 'CENTER': 'center', 'RIGHT': 'right', 'JUSTIFY': 'justify'}

IMAGE_POSITION = {'center': 'center', 'middle': 'center', 'left': 'left', 'right': 'right', 'top': 'top', 'bottom': 'bottom'}

WRAP_STRATEGY_MAP = {'OVERFLOW': 'no-wrap', 'CLIP': 'no-wrap', 'WRAP': 'wrap'}

COLUMNS = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
            'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
            'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ']
