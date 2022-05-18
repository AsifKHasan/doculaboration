#!/usr/bin/env python3

''' various utilities for generating an Openoffice odt document
'''
import platform
import subprocess
import random
import string
from pathlib import Path
from odf import style, text, draw, table
from helper.logger import *


if platform.system() == 'Windows':
    LIBREOFFICE_EXECUTABLE = 'C:/Program Files/LibreOffice/program/soffice.exe'
else:
    LIBREOFFICE_EXECUTABLE = 'soffice'


# --------------------------------------------------------------------------------------------------------------------------------------------
# pictures, background image

''' graphic-style
    <style:style style:name="fr1" style:family="graphic" style:parent-style-name="Graphics">
      <style:graphic-properties style:wrap="none" style:vertical-pos="top" style:vertical-rel="paragraph" style:horizontal-pos="center" style:horizontal-rel="page" style:mirror="none" fo:clip="rect(0in, 0in, 0in, 0in)" draw:luminance="0%" draw:contrast="0%" draw:red="0%" draw:green="0%" draw:blue="0%" draw:gamma="100%" draw:color-inversion="false" draw:image-opacity="100%" draw:color-mode="standard" draw:wrap-influence-on-position="once-concurrent" loext:allow-overlap="true"/>
    </style:style>
'''
def create_graphic_style(odt, valign, halign):
    style_name = f"fr-{random_string()}"

    graphic_properties_attributes = {'wrap': 'none', 'verticalpos': valign, 'horizontalpos': halign}
    graphic_properties = style.GraphicProperties(attributes=graphic_properties_attributes)

    graphic_style_attributes = {'name': style_name, 'family': 'graphic', 'parentstylename': 'Graphics'}
    graphic_style = style.Style(attributes=graphic_style_attributes)

    graphic_style.addElement(graphic_properties)
    odt.automaticstyles.addElement(graphic_style)

    return style_name


''' frame and image
    <draw:frame draw:style-name="fr1" draw:name="Image1" text:anchor-type="paragraph" svg:width="1.5in" svg:height="1.9in" draw:z-index="0">
      <draw:image xlink:href="Pictures/1000000000000258000002F8CC673C705E8CE146.jpg" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="image/jpeg"/>
    </draw:frame>
'''
def create_image_frame(odt, picture_path, valign, halign, width, height):
    # THIS IS THE Draw:Frame object to return
    draw_frame = None

    # first the image to be added into the document
    href = odt.addPicture(picture_path)
    if href:
        # next we need the Draw:Image object
        image_attributes = {'href': href}
        # image_attributes[('draw', 'mimetype')] = 'image/png'
        draw_image = draw.Image(attributes=image_attributes)

        frame_style_name = create_graphic_style(odt, valign, halign)

        # finally we need the Draw:Frame object
        frame_attributes = {'stylename': frame_style_name, 'anchortype': 'paragraph', 'width': f"{width}in", 'height': f"{height}in"}
        draw_frame = draw.Frame(attributes=frame_attributes)

        draw_frame.addElement(draw_image)

    else:
        warn(f"image {picture_path} copuld not be added into the document")

    return draw_frame



# --------------------------------------------------------------------------------------------------------------------------------------------
# table, table-row, table-column, table-cell

''' create a Table
'''
def create_table(odt, table_name, table_style_attributes, table_properties_attributes):
    if 'family' not in table_style_attributes:
        table_style_attributes['family'] = 'table'

    # create the style
    table_style = style.Style(attributes=table_style_attributes)
    table_style.addElement(style.TableProperties(attributes=table_properties_attributes))
    odt.automaticstyles.addElement(table_style)

    # create the table
    table_properties = {'name': table_name, 'stylename': table_style_attributes['name']}
    tbl = table.Table(attributes=table_properties)

    return tbl


''' create table-header-rows
'''
def create_table_header_rows():
    return table.TableHeaderRows()


''' create TableColumn
'''
def create_table_column(odt, table_column_name, table_column_style_attributes, table_column_properties_attributes):
    if 'family' not in table_column_style_attributes:
        table_column_style_attributes['family'] = 'table-column'

    # create the style
    table_column_style = style.Style(attributes=table_column_style_attributes)
    table_column_style.addElement(style.TableColumnProperties(attributes=table_column_properties_attributes))
    odt.automaticstyles.addElement(table_column_style)

    # create the table-column
    table_column_properties = {'stylename': table_column_style_attributes['name']}
    table_column = table.TableColumn(attributes=table_column_properties)

    return table_column


''' create TableRow
'''
def create_table_row(odt, table_row_style_attributes, table_row_properties_attributes):
    if 'family' not in table_row_style_attributes:
        table_row_style_attributes['family'] = 'table-row'

    # create the style
    table_row_style = style.Style(attributes=table_row_style_attributes)
    table_row_properties_attributes['keeptogether'] = 'always'
    table_row_style.addElement(style.TableRowProperties(attributes=table_row_properties_attributes))
    odt.automaticstyles.addElement(table_row_style)

    # create the table-row
    table_row_properties = {'stylename': table_row_style_attributes['name']}
    table_row = table.TableRow(attributes=table_row_properties)

    return table_row


''' create TableCell
'''
def create_table_cell(odt, table_cell_style_attributes, table_cell_properties_attributes, table_cell_attributes, background_image_style=None):
    if 'family' not in table_cell_style_attributes:
        table_cell_style_attributes['family'] = 'table-cell'

    # create the style
    table_cell_style = style.Style(attributes=table_cell_style_attributes)
    table_cell_properties = style.TableCellProperties(attributes=table_cell_properties_attributes)

    if background_image_style:
        table_cell_properties.addElement(background_image_style)

    table_cell_style.addElement(table_cell_properties)
    odt.automaticstyles.addElement(table_cell_style)

    # create the table-cell
    table_cell_attributes['stylename'] = table_cell_style_attributes['name']
    table_cell = table.TableCell(attributes=table_cell_attributes)

    return table_cell


''' create CoveredTableCell
'''
def create_covered_table_cell(odt, table_cell_style_attributes, table_cell_properties_attributes):
    if 'family' not in table_cell_style_attributes:
        table_cell_style_attributes['family'] = 'table-cell'

    # create the style
    table_cell_style = style.Style(attributes=table_cell_style_attributes)
    table_cell_style.addElement(style.TableCellProperties(attributes=table_cell_properties_attributes))
    odt.automaticstyles.addElement(table_cell_style)

    # create the table-cell
    table_cell_attributes = {'stylename': table_cell_style_attributes['name']}
    table_cell = table.CoveredTableCell(attributes=table_cell_attributes)

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
    paragraph_style = style.Style(attributes=style_attributes)

    if paragraph_attributes is not None:
        paragraph_style.addElement(style.ParagraphProperties(attributes=paragraph_attributes))

    if text_attributes is not None:
        paragraph_style.addElement(style.TextProperties(attributes=text_attributes))

    odt.automaticstyles.addElement(paragraph_style)

    return paragraph_style.getAttribute('name')


''' page number
    <text:p text:style-name="MP1">Page <text:page-number text:select-page="current">1</text:page-number>
        <text:s/>of <text:page-count>1</text:page-count>
    </text:p>
'''
def create_page_number(style_name, short=False):
    paragraph = text.P(stylename=style_name)
    page_number = text.PageNumber(selectpage='current')

    if short:
        paragraph.addText("Page ")
        paragraph.addElement(page_number)
    else:
        s = text.S()
        page_count = text.PageCount()

        paragraph.addText("Page ")
        paragraph.addElement(page_number)
        paragraph.addElement(s)
        paragraph.addText("of ")
        paragraph.addElement(page_count)

    return paragraph


''' write a paragraph in a given style
'''
def create_paragraph(odt, style_name, text_content=None, run_list=None, outline_level=0):
    style = odt.getStyleByName(style_name)
    if style is None:
        warn(f"style {style_name} not found")

    if run_list is not None:
        paragraph = text.P(stylename=style)
        for run in run_list:
            style_attributes = {'family': 'text'}
            text_style_name = create_paragraph_style(odt, style_attributes=style_attributes, text_attributes=run['text-attributes'])
            paragraph.addElement(text.Span(stylename=text_style_name, text=run['text']))

        return paragraph

    if text_content is not None:
        if outline_level == 0:
            paragraph = text.P(stylename=style_name, text=text_content)
        else:
            paragraph = text.H(stylename=style_name, text=text_content, outlinelevel=outline_level)

        return paragraph
    else:
        paragraph = text.P(stylename=style_name)
        return paragraph



# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# indexes and pdf generation

''' update indexes through a macro macro:///Standard.Module1.open_document(document_url) which must be in OpenOffice macro library
'''
def update_indexes(odt, odt_path):
    debug(msg=f"updating indexes for {Path(odt_path).resolve()}")
    document_url = Path(odt_path).as_uri()
    macro = f'"macro:///Standard.Module1.open_document("{document_url}")"'
    command_line = f'"{LIBREOFFICE_EXECUTABLE}" --headless --invisible {macro}'
    subprocess.call(command_line, shell=True);


''' given an odt file generates pdf in the given directory
'''
def generate_pdf(infile, outdir):
    debug(msg=f"generating pdf from {Path(infile).resolve()}")
    command_line = f'"{LIBREOFFICE_EXECUTABLE}" --headless --convert-to pdf --outdir "{outdir}" "{infile}"'
    subprocess.call(command_line, shell=True);


''' create table-of-contents
  <text:table-of-content-entry-template text:outline-level="1" text:style-name="Contents_20_1">
    <text:index-entry-link-start text:style-name="Index_20_Link"/>
    <text:index-entry-chapter/>
    <text:index-entry-text/>
    <text:index-entry-tab-stop style:type="right" style:leader-char="."/>
    <text:index-entry-page-number/>
    <text:index-entry-link-end/>
  </text:table-of-content-entry-template>
'''
def create_toc():
    name = 'Table of Content'
    toc = text.TableOfContent(name=name)
    toc_source = text.TableOfContentSource(outlinelevel=10)
    toc_title_template = text.IndexTitleTemplate()
    toc_source.addElement(toc_title_template)

    for i in range(1, 11):
        toc_entry_template = text.TableOfContentEntryTemplate(outlinelevel=i, stylename=f"Contents_20_{i}")

        index_entry_link_start = text.IndexEntryLinkStart(stylename='Index_20_Link')
        index_entry_chapter = text.IndexEntryChapter()
        index_entry_text = text.IndexEntryText()
        index_entry_tab_stop = text.IndexEntryTabStop(type='right', leaderchar='.')
        index_entry_page_number = text.IndexEntryPageNumber()
        index_entry_link_end = text.IndexEntryLinkEnd()

        toc_entry_template.addElement(index_entry_link_start)
        toc_entry_template.addElement(index_entry_chapter)
        toc_entry_template.addElement(index_entry_text)
        toc_entry_template.addElement(index_entry_tab_stop)
        toc_entry_template.addElement(index_entry_page_number)
        toc_entry_template.addElement(index_entry_link_end)

        toc_source.addElement(toc_entry_template)

    toc.addElement(toc_source)

    return toc


''' create illustration-index
'''
def create_lof():
    name = 'List of Figures'
    toc = text.TableOfContent(name=name)
    toc_source = text.TableOfContentSource(outlinelevel=1, useoutlinelevel=False, useindexmarks=False, useindexsourcestyles=True)
    toc_entry_template = text.TableOfContentEntryTemplate(outlinelevel=1, stylename='Figure')

    index_entry_link_start = text.IndexEntryLinkStart(stylename='Index_20_Link')
    index_entry_chapter = text.IndexEntryChapter()
    index_entry_text = text.IndexEntryText()
    index_entry_tab_stop = text.IndexEntryTabStop(type='right', leaderchar='.')
    index_entry_page_number = text.IndexEntryPageNumber()
    index_entry_link_end = text.IndexEntryLinkEnd()

    toc_entry_template.addElement(index_entry_link_start)
    toc_entry_template.addElement(index_entry_chapter)
    toc_entry_template.addElement(index_entry_text)
    toc_entry_template.addElement(index_entry_tab_stop)
    toc_entry_template.addElement(index_entry_page_number)
    toc_entry_template.addElement(index_entry_link_end)

    toc_title_template = text.IndexTitleTemplate()
    toc_index_source_styles = text.IndexSourceStyles(outlinelevel=1)
    toc_index_source_style = text.IndexSourceStyle(stylename='Figure')
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
    toc = text.TableOfContent(name=name)
    toc_source = text.TableOfContentSource(outlinelevel=1, useoutlinelevel=False, useindexmarks=False, useindexsourcestyles=True)
    toc_entry_template = text.TableOfContentEntryTemplate(outlinelevel=1, stylename='Table')

    index_entry_link_start = text.IndexEntryLinkStart(stylename='Index_20_Link')
    index_entry_chapter = text.IndexEntryChapter()
    index_entry_text = text.IndexEntryText()
    index_entry_tab_stop = text.IndexEntryTabStop(type='right', leaderchar='.')
    index_entry_page_number = text.IndexEntryPageNumber()
    index_entry_link_end = text.IndexEntryLinkEnd()

    toc_entry_template.addElement(index_entry_link_start)
    toc_entry_template.addElement(index_entry_chapter)
    toc_entry_template.addElement(index_entry_text)
    toc_entry_template.addElement(index_entry_tab_stop)
    toc_entry_template.addElement(index_entry_page_number)
    toc_entry_template.addElement(index_entry_link_end)

    toc_title_template = text.IndexTitleTemplate()
    toc_index_source_styles = text.IndexSourceStyles(outlinelevel=1)
    toc_index_source_style = text.IndexSourceStyle(stylename='Table')
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
    for master_page in odt.masterstyles.getElementsByType(style.MasterPage):
        if master_page.getAttribute('name') == master_page_name:
            return master_page

    warn(f"master-page {master_page_name} NOT found")
    return None


''' get page-layout by name
'''
def get_page_layout(odt, page_layout_name):
    for page_layout in odt.automaticstyles.getElementsByType(style.PageLayout):
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
    page_layout = style.PageLayout(name=page_layout_name)
    odt.automaticstyles.addElement(page_layout)

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
    page_layout.addElement(style.PageLayoutProperties(pagewidth=pagewidth, pageheight=pageheight, margintop=margintop, marginbottom=marginbottom, marginleft=marginleft, marginright=marginright, printorientation=orientation))

    return page_layout


''' create (section-specific) master-page
    page layouts are saved with a name mp-section-no
'''
def create_master_page(odt, odt_specs, master_page_name, page_layout_name, page_spec, margin_spec, orientation):
    # create one, first get/create the page-layout
    page_layout = create_page_layout(odt, odt_specs, page_layout_name, page_spec, margin_spec, orientation)
    master_page = style.MasterPage(name=master_page_name, pagelayoutname=page_layout_name)
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
        height = f"{HEADER_HEIGHT}in"
        if odd_even == 'odd':
            header_footer_style = style.Header()
        if odd_even == 'even':
            header_footer_style = style.HeaderLeft()
        if odd_even == 'first':
            # header_footer_style = HeaderFirst()
            pass

    elif header_footer == 'footer':
        height = f"{FOOTER_HEIGHT}in"
        if odd_even == 'odd':
            header_footer_style = style.Footer()
        if odd_even == 'even':
            header_footer_style = style.FooterLeft()
        if odd_even == 'first':
            # header_footer_style = FooterFirst()
            pass

    if header_footer_style:
        # TODO: the height should come from actual header content height
        header_footer_properties_attributes = {'margin': '0in', 'padding': '0in', 'dynamicspacing': False, 'height': height}
        header_style = style.HeaderStyle()
        header_style.addElement(style.HeaderFooterProperties(attributes=header_footer_properties_attributes))
        footer_style = style.FooterStyle()
        footer_style.addElement(style.HeaderFooterProperties(attributes=header_footer_properties_attributes))
        page_layout.addElement(header_style)
        page_layout.addElement(footer_style)

        master_page.addElement(header_footer_style)

    return header_footer_style



# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# various utility functions

''' whether the container is a table-cell or not
'''
def is_table_cell(container):
    # if container is n instance of table-cell
    if isinstance(container, type(table.TableCell)):
        return True
    else:
        return False


''' given pixel size, calculate the row height in inches
    a reasonable approximation is what gsheet says 21 pixels, renders well as 12 pixel (assuming our normal text is 10-11 in size)
'''
def row_height_in_inches(pixel_size):
    return float((pixel_size) / 96)


''' get a random string
'''
def random_string(length=12):
    letters = string.ascii_uppercase
    return ''.join(random.choice(letters) for i in range(length))


''' fit width/height into a given width/height maintaining aspect ratio
'''
def fit_width_height(fit_within_width, fit_within_height, width_to_fit, height_to_fit):
    WIDTH_OFFSET = 0.0
    HEIGHT_OFFSET = 0.2

    fit_within_width = fit_within_width - WIDTH_OFFSET
    fit_within_height = fit_within_height - HEIGHT_OFFSET

    aspect_ratio = width_to_fit / height_to_fit

    if width_to_fit > fit_within_width:
        width_to_fit = fit_within_width
        height_to_fit = width_to_fit / aspect_ratio
        if height_to_fit > fit_within_height:
            height_to_fit = fit_within_height
            width_to_fit = height_to_fit * aspect_ratio

    return width_to_fit, height_to_fit



# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# various utility data

COLSEP = (0/72)
# ROWSEP = (2/72)

HEADER_HEIGHT = 0.3
FOOTER_HEIGHT = 0.3

GSHEET_ODT_BORDER_MAPPING = {
    'DOTTED': 'dotted',
    'DASHED': 'dash',
    'SOLID': 'solid'
}


HEADING_TO_LEVEL = {
    'Heading 1': {'outline-level': 1},
    'Heading 2': {'outline-level': 2},
    'Heading 3': {'outline-level': 3},
    'Heading 4': {'outline-level': 4},
    'Heading 5': {'outline-level': 5},
    'Heading 6': {'outline-level': 6},
    'Heading 7': {'outline-level': 7},
    'Heading 8': {'outline-level': 8},
    'Heading 9': {'outline-level': 9},
    'Heading 10': {'outline-level': 10},
}


LEVEL_TO_HEADING = [
    'Title',
    'Heading 1',
    'Heading 2',
    'Heading 3',
    'Heading 4',
    'Heading 5',
    'Heading 6',
    'Heading 7',
    'Heading 8',
    'Heading 9',
    'Heading 10',
]


TEXT_VALIGN_MAP = {'TOP': 'top', 'MIDDLE': 'middle', 'BOTTOM': 'bottom'}
# TEXT_HALIGN_MAP = {'LEFT': 'start', 'CENTER': 'center', 'RIGHT': 'end', 'JUSTIFY': 'justify'}
TEXT_HALIGN_MAP = {'LEFT': 'left', 'CENTER': 'center', 'RIGHT': 'right', 'JUSTIFY': 'justify'}

IMAGE_POSITION = {'center': 'center', 'middle': 'center', 'left': 'left', 'right': 'right', 'top': 'top', 'bottom': 'bottom'}

WRAP_STRATEGY_MAP = {'OVERFLOW': 'no-wrap', 'CLIP': 'no-wrap', 'WRAP': 'wrap'}

COLUMNS = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
            'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
            'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ']
