#!/usr/bin/env python

''' various utilities for generating an Openoffice odt document
'''
import re
import platform
import subprocess
import random
import string
import importlib

from pathlib import Path
from odf import style, text, draw, table
from odf.style import Style, TextProperties, ParagraphProperties, PageLayout, MasterPage, Header, HeaderLeft, Footer, FooterLeft, FontFace
from odf.element import Element
from namespaces import MATHNS
from odf.namespaces import DRAWNS

from xml.dom.minidom import parseString
from xml.dom import Node

import latex2mathml.converter

from helper.logger import *


if platform.system() == 'Windows':
    LIBREOFFICE_EXECUTABLE = 'C:/Program Files/LibreOffice/program/soffice.exe'
else:
    LIBREOFFICE_EXECUTABLE = 'soffice'


''' process a list of section_data and generate odt code
'''
def section_list_to_odt(section_list, config, nesting_level=0):
    first_section = True
    for section in section_list:
        section_meta = section['section-meta']
        section_prop = section['section-prop']

        if section_prop['label'] != '':
            info(f"writing : {section_prop['label'].strip()} {section_prop['heading'].strip()}", nesting_level=section_meta['nesting-level'])
        else:
            info(f"writing : {section_prop['heading'].strip()}", nesting_level=section_meta['nesting-level'])


        section_meta['first-section'] = first_section
        if first_section:
            first_section = False


        module = importlib.import_module("odt.odt_api")
        func = getattr(module, f"process_{section_prop['content-type']}")
        func(section, config)


# --------------------------------------------------------------------------------------------------------------------------------------------
# pictures, background image

''' background-image style
    <style:background-image
        xlink:href="Pictures/10000001000007D0000007D0EF304D419C6347C7.png"
        xlink:type="simple" xlink:actuate="onLoad" />

    style:page-layout -> style:page-layout-properties
        draw:fill-image-width="0cm"
        draw:fill-image-height="0cm"
        draw:fill-image-ref-point-x="0%"
        draw:fill-image-ref-point-y="0%"
        draw:fill-image-ref-point="center"
        draw:tile-repeat-offset="0% vertical"
'''
def create_background_image_style(odt, picture_path, nesting_level=0):
    background_image_style = None

    # first the image to be added into the document
    href = odt.addPicture(picture_path)
    if href:
        background_image_style_attributes = {'href': href, 'opacity': '100%', 'position': 'center center', 'repeat': 'stretch', }

        # background_image_style_attributes = {'href': href}
        background_image_style = style.BackgroundImage(attributes=background_image_style_attributes)

    else:
        warn(f"image {picture_path} could not be added into the document", nesting_level=nesting_level)

    return background_image_style


''' graphic-style
    <style:style style:name="fr1" style:family="graphic" style:parent-style-name="Graphics">
      <style:graphic-properties style:wrap="none" style:vertical-pos="top" style:vertical-rel="paragraph" style:horizontal-pos="center" style:horizontal-rel="page" style:mirror="none" fo:clip="rect(0in, 0in, 0in, 0in)" draw:luminance="0%" draw:contrast="0%" draw:red="0%" draw:green="0%" draw:blue="0%" draw:gamma="100%" draw:color-inversion="false" draw:image-opacity="100%" draw:color-mode="standard" draw:wrap-influence-on-position="once-concurrent" loext:allow-overlap="true"/>
    </style:style>
'''
def create_graphic_style(odt, graphic_properties_attributes, nesting_level=0):
    style_name = f"fr-{random_string()}"

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
def create_image_frame(odt, picture_path, graphic_properties_attributes, frame_attributes, nesting_level=0):
    draw_frame = None

    # first the image to be added into the document
    href = odt.addPicture(picture_path)
    if href:
        # next we need the Draw:Image object
        image_attributes = {'href': href}
        draw_image = draw.Image(attributes=image_attributes)

        # create the graphic-style
        frame_style_name = create_graphic_style(odt=odt, graphic_properties_attributes=graphic_properties_attributes)

        # finally we need the Draw:Frame object
        frame_attributes = {**{'stylename': frame_style_name}, **frame_attributes}
        draw_frame = draw.Frame(attributes=frame_attributes)
        draw_frame.addElement(draw_image)

    else:
        warn(f"image {picture_path} could not be added into the document")

    return draw_frame


''' create a text_a
'''
def create_text_a(anchor, target, nesting_level=0):
    
    # if the anchor is not an url, it is a bookmark
    if not target.startswith('http'):
        target_text = f"#{target.strip()}"
        if anchor:
            anchor_text = anchor
        else:
            # TODO: it should actually be the text associated with the target bookmark
            anchor_text = target_text
    else:
        target_text = target.strip()
        if anchor:
            anchor_text = anchor
        else:
            anchor_text = target_text

    link = text.A(text=anchor_text, type="simple", href=target_text)

    return link


# --------------------------------------------------------------------------------------------------------------------------------------------
# table, table-row, table-column, table-cell

''' create a Table
'''
def create_table(odt, table_name, table_style_attributes, table_properties_attributes, nesting_level=0):
    if 'family' not in table_style_attributes:
        table_style_attributes['family'] = 'table'

    # table_style['border-model'] = 'collapsing'

    # create the style
    table_style = style.Style(attributes=table_style_attributes)
    table_style.addElement(style.TableProperties(attributes=table_properties_attributes, bordermodel='collapsing'))
    odt.automaticstyles.addElement(table_style)

    # create the table
    table_properties = {'name': table_name, 'stylename': table_style_attributes['name']}
    tbl = table.Table(attributes=table_properties)

    return tbl


''' create table-header-rows
'''
def create_table_header_rows(nesting_level=0):
    return table.TableHeaderRows()


''' create TableColumn
'''
def create_table_column(odt, table_column_name, table_column_style_attributes, table_column_properties_attributes, nesting_level=0):
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
def create_table_row(odt, table_row_style_attributes, table_row_properties_attributes, nesting_level=0):
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
def create_table_cell(odt, table_cell_style_attributes, table_cell_properties_attributes, table_cell_attributes, background_image_style=None, nesting_level=0):
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
def create_covered_table_cell(odt, table_cell_style_attributes, table_cell_properties_attributes, nesting_level=0):
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

''' create a Paragraph Style that forces a switch to this Master Page
    This is how we change backgrounds mid-document
'''
def create_paragraph_with_masterpage(odt, style_name, master_page_name, nesting_level=0):
    ps = style.Style(name=style_name, family="paragraph", masterpagename=master_page_name)
    odt.automaticstyles.addElement(ps)
    p = text.P(stylename=style_name, text=None)

    return p


''' create style - family paragraph
'''
def create_paragraph_style(odt, style_attributes=None, paragraph_attributes=None, text_attributes=None, nesting_level=0):
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
    paragraph_style = style.Style(name=style_attributes['name'], attributes=style_attributes)

    if paragraph_attributes is not None:
        paragraph_style.addElement(style.ParagraphProperties(attributes=paragraph_attributes))

    if text_attributes is not None:
        paragraph_style.addElement(style.TextProperties(attributes=text_attributes))

    odt.automaticstyles.addElement(paragraph_style)

    return style_attributes['name']


''' write a paragraph in a given style
'''
def create_paragraph(odt, style_name, text_content=None, run_list=None, outline_level=0, footnote_list={}, bookmark={}, keep_line_breaks=False, directives=True, nesting_level=0):
    style = odt.getStyleByName(style_name)
    if style is None:
        warn(f"style {style_name} not found")

    paragraph = None

    # text-runs
    if run_list is not None:
        if outline_level == 0:
            paragraph = text.P(stylename=style)
        else:
            paragraph = text.H(stylename=style, outlinelevel=outline_level)

        for run in run_list:
            style_attributes = {'family': 'text'}
            text_style_name = create_paragraph_style(odt, style_attributes=style_attributes, text_attributes=run['text-attributes'])
            fragment = create_text(text_type='span', style_name=text_style_name, text_content=run['text'], footnote_list=footnote_list, bookmark=bookmark, keep_line_breaks=keep_line_breaks)
            paragraph.addElement(fragment)

    # P or H
    elif text_content is not None:
        if outline_level == 0:
            paragraph = create_text(text_type='P', style_name=style_name, text_content=text_content, footnote_list=footnote_list, bookmark=bookmark, keep_line_breaks=keep_line_breaks)
        else:
            paragraph = create_text(text_type='H', style_name=style_name, text_content=text_content, outline_level=outline_level, footnote_list=footnote_list, bookmark=bookmark, keep_line_breaks=keep_line_breaks)

    else:
        paragraph = text.P(stylename=style_name)


    return paragraph


''' add text span, s, c to a paragraph
'''
def add_text_to_paragraph(paragraph, text_string, nesting_level=0):
    matches = []
    matches = matches + [(match.start(), match.end(), 'space') for match in re.finditer(r'^ {1,}', text_string)]
    matches = matches + [(match.start(), match.end(), 'space') for match in re.finditer(r' {2,}', text_string)]
    matches = matches + [(match.start(), match.end(), 'space') for match in re.finditer(r' {1,}$', text_string)]

    non_matches = []
    non_matches = non_matches + [(match.start(), match.end(), 'text') for match in re.finditer(r'[^ ]+(?: [^ ]+)*', text_string)]
    all = matches + non_matches
    all.sort()

    for start, end, what in all:
        if what == 'space':
            paragraph.addElement(text.S(c=end-start))

        elif what == 'text':
            paragraph.addText(text=text_string[start:end])


''' create a P or H or span
'''
def create_text(text_type, style_name, text_content=None, outline_level=0, footnote_list={}, bookmark={}, keep_line_breaks=False, nesting_level=0):
    paragraph = None

    # process FN{...} first, we get a list of block dicts
    inline_blocks = process_footnotes(text_content=text_content, footnote_list=footnote_list)

    # process LATEX$...$ for each text item
    new_inline_blocks = []
    for inline_block in inline_blocks:
        # process only 'text'
        if 'text' in inline_block:
            new_inline_blocks = new_inline_blocks + process_latex_blocks(inline_block['text'])

        else:
            new_inline_blocks.append(inline_block)

    inline_blocks = new_inline_blocks

    # process PAGE{..} for each text item
    new_inline_blocks = []
    for inline_block in inline_blocks:
        # process only 'text'
        if 'text' in inline_block:
            new_inline_blocks = new_inline_blocks + process_bookmark_page_blocks(inline_block['text'])

        else:
            new_inline_blocks.append(inline_block)

    inline_blocks = new_inline_blocks

    # process LINK{..} for each text item
    new_inline_blocks = []
    for inline_block in inline_blocks:
        # process only 'text'
        if 'text' in inline_block:
            new_inline_blocks = new_inline_blocks + process_links(inline_block['text'])

        else:
            new_inline_blocks.append(inline_block)

    inline_blocks = new_inline_blocks

    # create the P or H or span
    if text_type == 'P':
        paragraph = text.P(stylename=style_name)

    elif text_type == 'H':
        paragraph = text.H(stylename=style_name, outlinelevel=outline_level)

    elif text_type == 'span':
        paragraph = text.Span(stylename=style_name)

    # bookmark
    if bookmark:
        for k, v in bookmark.items():
            # in odt bookmarks only has name, no text
            paragraph.addElement(text.Bookmark(name=k))

    # now fill the paragraph with texts and footnotes
    for inline_block in inline_blocks:
        if 'text' in inline_block:
            if keep_line_breaks:
                splits = inline_block['text'].split('\n')
                if len(splits) == 1:
                    add_text_to_paragraph(paragraph=paragraph, text_string=inline_block['text'])

                else:
                    add_text_to_paragraph(paragraph=paragraph, text_string=splits[0])
                    for part in splits[1:]:
                        paragraph.addElement(text.LineBreak())
                        add_text_to_paragraph(paragraph=paragraph, text_string=part)

            else:
                add_text_to_paragraph(paragraph=paragraph, text_string=inline_block['text'])

        elif 'fn' in inline_block:
            footnote_object = create_footnote(inline_block['fn'])
            paragraph.addElement(footnote_object)

        elif 'latex' in inline_block:
            latex_df = create_latex(latex_content=inline_block['latex'])
            if latex_df:
                paragraph.addElement(latex_df)

        elif 'page-num' in inline_block:
            page_num = text.PageNumber(selectpage='current')
            paragraph.addElement(page_num)

        elif 'page-count' in inline_block:
            page_count = text.PageCount()
            paragraph.addElement(page_count)

        elif 'bookmark-page' in inline_block:
            bookmark_ref = inline_block['bookmark-page'].strip()
            if bookmark_ref != '':
                paragraph.addElement(text.BookmarkRef(refname=bookmark_ref, referenceformat='page'))

        elif 'link' in inline_block:
            target, anchor = inline_block['link'][0], inline_block['link'][1]
            # target is an xlink
            text_a = create_text_a(anchor=anchor, target=target)
            paragraph.addElement(text_a)

    return paragraph


''' process footnotes inside text
'''
def process_footnotes(text_content, footnote_list, nesting_level=0):
    # if text contains footnotes we make a list containing texts->footnote->text->footnote ......
    texts_and_footnotes = []

    # find out if there is any match with FN{key} inside the text_content
    pattern = r'FN{[^}]+}'
    current_index = 0
    for match in re.finditer(pattern, text_content):
      footnote_key = match.group()[3:-1]
      if footnote_key in footnote_list:
        # debug(f".... footnote {footnote_key} found at {match.span()} with description")
        # we have found a footnote, we add the preceding text and the footnote spec into the list
        footnote_start_index, footnote_end_index = match.span()[0], match.span()[1]
        if footnote_start_index >= current_index:
          # there are preceding text before the footnote
          texts_and_footnotes.append({'text': text_content[current_index:footnote_start_index]})
          texts_and_footnotes.append({'fn': (footnote_key, footnote_list[footnote_key])})
          current_index = footnote_end_index
      else:
        warn(f".... footnote {footnote_key} found at {match.span()}, but no details found")
        # this is not a footnote, ignore it
        footnote_start_index, footnote_end_index = match.span()[0], match.span()[1]
        # current_index = footnote_end_index + 1

    # there may be trailing text
    texts_and_footnotes.append({'text': text_content[current_index:]})

    return texts_and_footnotes


''' process latex blocks inside text
'''
def process_latex_blocks(text_content, nesting_level=0):
    # find out if there is any match with LATEX$...$ inside the text_content
    texts_and_latex = []

    pattern = r'LATEX\$[^$]+\$'
    current_index = 0
    for match in re.finditer(pattern, text_content):
        latex_content = match.group()[6:-1]

        # we have found a latex block, we add the preceding text and the latex block into the list
        latex_start_index, latex_end_index = match.span()[0], match.span()[1]
        if latex_start_index >= current_index:
            # there are preceding text before the latex
            text = text_content[current_index:latex_start_index]

            texts_and_latex.append({'text': text})

            texts_and_latex.append({'latex': latex_content})

            current_index = latex_end_index


    # there may be trailing text
    text = text_content[current_index:]

    texts_and_latex.append({'text': text})

    return texts_and_latex


''' process bookmark page ref inside text
'''
def process_bookmark_page_blocks(text_content, nesting_level=0):
    # find out if there is any match with PAGE{...} inside the text_content
    texts_and_bookmarks = []

    pattern = r'PAGE{[^}]*}'
    current_index = 0
    for match in re.finditer(pattern, text_content):
        bookmark_content = match.group()[5:-1].strip()

        # we have found a bookmark block, we add the preceding text and the bookmark block into the list
        bookmark_start_index, bookmark_end_index = match.span()[0], match.span()[1]
        if bookmark_start_index >= current_index:
            # there are preceding text before the bookmark
            text = text_content[current_index:bookmark_start_index]

            texts_and_bookmarks.append({'text': text})

            # PAGE{} means current page, PAGE{*} means number of pages, PAGE{XYZ} means page number where bookmark XYZ is set
            if bookmark_content == '':
                texts_and_bookmarks.append({"page-num": None})

            elif bookmark_content == '*':
                texts_and_bookmarks.append({'page-count': None})

            else:
                texts_and_bookmarks.append({'bookmark-page': bookmark_content})

            current_index = bookmark_end_index

    # there may be trailing text
    text = text_content[current_index:]

    texts_and_bookmarks.append({'text': text})

    return texts_and_bookmarks


''' process external url or bookmark links
'''
def process_links(text_content, nesting_level=0):
    # find out if there is any match with LINK{...}{} inside the text_content
    links = []

    # LINK patterns are like LINK{target}{text} or LINK{target}
    pattern = r'LINK({[^}]*}){1,2}'
    current_index = 0
    for match in re.finditer(pattern, text_content):
        link_content_pattern = r'([^{}]+)'
        i = 0
        target, anchor = None, None
        for content_match in re.finditer(link_content_pattern, match.group()):
            if i == 1:
                target = content_match.group()
            
            if i == 2:
                anchor = content_match.group()

            i = i + 1

        # we have found a LINK block, we add the preceding text and the LINK block into the list
        link_start_index, link_end_index = match.span()[0], match.span()[1]
        if link_start_index >= current_index:
            # there are preceding text before the link
            text = text_content[current_index:link_start_index]
            links.append({'text': text})

            # LINK patterns are like LINK{target}{text} or LINK{target}
            if target:
                links.append({"link": [target, anchor]})

            current_index = link_end_index

    # there may be trailing text
    text = text_content[current_index:]

    links.append({'text': text})

    return links


''' create a footnote
'''
def create_footnote(footnote_tuple, nesting_level=0):
    citation, footnote = footnote_tuple[0], footnote_tuple[1]
    note = text.Note(noteclass='footnote')
    note_citation = text.NoteCitation(label=citation)
    note_body = text.NoteBody()
    patagraph = text.P(text=footnote)
    note_body.addElement(patagraph)
    note.addElement(note_citation)
    note.addElement(note_body)

    return note


''' create latex object
'''
def create_latex(latex_content, nesting_level=0):
    # convert to MathML
    if latex_content is not None:
        mathml_output = latex2mathml.converter.convert(strip_math_mode_delimeters(latex_content))
        draw_frame = draw.Frame(zindex=0, anchortype='as-char')

        math = mathml_odf(mathml_content=mathml_output)
        draw_object = draw.Object()
        draw_object.addElement(math)
        draw_frame.addElement(draw_object)

        return draw_frame

    else:
        return None


''' write a mathml draw-frame
'''
def create_mathml(odt, style_name, latex_content, nesting_level=0):
    # process styles
    style = odt.getStyleByName(style_name)
    if style is None:
        warn(f"style {style_name} not found")

    paragraph = text.P(stylename=style)

    latex_df = create_latex(latex_content=latex_content)
    if latex_df is not None:
        paragraph.addElement(latex_df)

    return paragraph


''' odf.math.Math element
'''
def mathml_odf(mathml_content, nesting_level=0):
    # TODO: process the generated MathML
    mathml = mathml_content
    math_ = parseString(mathml.encode('utf-8'))
    math_ = math_.documentElement
    odf_math = mathml_odf_(math_)

    return odf_math


''' odf.math.Math element generator
'''
def mathml_odf_(parent, nesting_level=0):
    elem = Element(qname = (MATHNS,parent.tagName))
    if parent.attributes:
        for attr, value in parent.attributes.items():
            elem.setAttribute((MATHNS,attr), value, check_grammar=False)

    for child in parent.childNodes:
        if child.nodeType == Node.TEXT_NODE:
            text = child.nodeValue
            elem.addText(text, check_grammar=False)
        else:
            elem.addElement(mathml_odf_(child), check_grammar=False)

    return elem


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# indexes and pdf generation

''' update indexes through a macro macro:///Standard.Module1.open_document(document_url) which must be in OpenOffice macro library
'''
def update_indexes(odt, odt_path, nesting_level=0):
    document_url = Path(odt_path).as_uri()

    macro = f'"macro:///Standard.Module1.force_update("{document_url}")"'
    command_line = f'"{LIBREOFFICE_EXECUTABLE}" --headless --invisible {macro}'
    subprocess.call(command_line, shell=True)

    macro = f'"macro:///Standard.Module1.open_document("{document_url}")"'
    command_line = f'"{LIBREOFFICE_EXECUTABLE}" --headless --invisible {macro}'
    subprocess.call(command_line, shell=True)


''' given an odt file generates pdf in the given directory
'''
def generate_pdf(infile, outdir, nesting_level=0):
    command_line = f'"{LIBREOFFICE_EXECUTABLE}" --headless --convert-to pdf:writer_pdf_Export --outdir "{outdir}" "{infile}"'
    subprocess.call(command_line, shell=True)

    #  now there should be a pdf file, we need to rename it
    file_to_rename = infile.replace(r'.odt', r'.pdf')
    rename_to = infile + '.pdf'
    Path(file_to_rename).replace(rename_to)


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
def create_toc(nesting_level=0):
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
def create_lof(nesting_level=0):
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
def create_lot(nesting_level=0):
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
def get_master_page(odt, master_page_name, nesting_level=0):
    for master_page in odt.masterstyles.getElementsByType(style.MasterPage):
        if master_page.getAttribute('name') == master_page_name:
            return master_page

    warn(f"master-page {master_page_name} NOT found")
    return None


''' get page-layout by name
'''
def get_page_layout(odt, page_layout_name, nesting_level=0):
    for page_layout in odt.automaticstyles.getElementsByType(style.PageLayout):
        if page_layout.getAttribute('name') == page_layout_name:
            return page_layout

    # warn(f"page-layout {page_layout_name} NOT found")
    return None


''' create (section-specific) page-layout
        fillimagewidth = '0cm'
        fillimageheight = '0cm'
        fillimagerefpointx = '0%'
        fillimagerefpointy = '0%'
        fillimagerefpoint = 'center'
        tilerepeatoffset = '0% vertical'
'''
def create_page_layout(odt, page_layout_name, page_spec, margin_spec, orientation, background_image_path, nesting_level=0):
    # get one, if not found create one
    page_layout = get_page_layout(odt=odt, page_layout_name=page_layout_name)
    if page_layout is None:
        page_layout = style.PageLayout(name=page_layout_name)
        odt.automaticstyles.addElement(page_layout)

    if orientation == 'portrait':
        pageheight = f"{page_spec['height']}in"
        pagewidth = f"{page_spec['width']}in"
    else:
        pageheight = f"{page_spec['width']}in"
        pagewidth = f"{page_spec['height']}in"

    margintop = f"{margin_spec['top']}in"
    marginbottom = f"{margin_spec['bottom']}in"
    marginleft = f"{margin_spec['left']}in"
    marginright = f"{margin_spec['right']}in"
    # margingutter = f"{margin_spec]['gutter']}in"

    page_layout_properties = None
    # background-image
    if background_image_path != '':
        background_image_style = create_background_image_style(odt, background_image_path)
        if background_image_style:

            # page_layout_properties = style.PageLayoutProperties(pagewidth=pagewidth, pageheight=pageheight, margintop=margintop, marginbottom=marginbottom, marginleft=marginleft, marginright=marginright, printorientation=orientation, fillimagewidth=fillimagewidth, fillimageheight=fillimageheight, fillimagerefpointx=fillimagerefpointx, fillimagerefpointy=fillimagerefpointy, fillimagerefpoint=fillimagerefpoint, tilerepeatoffset=tilerepeatoffset)
            page_layout_properties = style.PageLayoutProperties(pagewidth=pagewidth, pageheight=pageheight,
                                    margintop=margintop, marginbottom=marginbottom, marginleft=marginleft, marginright=marginright, printorientation=orientation
                                    )
            page_layout_properties.addElement(background_image_style)

            # # background_image specific page-layout-properties
            page_layout_attrs = {'fill-image-width': '100%', 'fill-image-height': '100%', 'fill-image-ref-point-x': '0%', 'fill-image-ref-point-y': '0%', 'fill-image-ref-point': 'center', 'tile-repeat-offset': '0% vertical'}

            for attr_name, attr_value in page_layout_attrs.items():
                page_layout_properties.setAttrNS(DRAWNS, attr_name, attr_value)

    if not page_layout_properties:
        page_layout_properties = style.PageLayoutProperties(pagewidth=pagewidth, pageheight=pageheight, margintop=margintop, marginbottom=marginbottom, marginleft=marginleft, marginright=marginright, printorientation=orientation)

    page_layout.addElement(page_layout_properties)

    return page_layout


''' create (section-specific) master-page
    page layouts are saved with a name mp-section-no
'''
def create_master_page(odt, first_section, document_index, master_page_name, page_layout_name, page_spec, margin_spec, orientation, background_image_path, next_master_page_style=None, nesting_level=0):
    # TODO: create one, first get/create the page-layout. If first section, update page-layout for *Standarad* master-page
    if first_section and document_index == 0:
        master_page = get_master_page(odt=odt, master_page_name='Standard')
        existing_page_layout_name = master_page.attributes[(master_page.qname[0], 'page-layout-name')]
        page_layout = create_page_layout(odt=odt, page_layout_name=existing_page_layout_name, page_spec=page_spec, margin_spec=margin_spec, orientation=orientation, background_image_path=background_image_path)
        # existing_page_layout = get_page_layout(odt=odt, page_layout_name=existing_page_layout_name)
    else:
        master_page = style.MasterPage(name=master_page_name, pagelayoutname=page_layout_name, nextstylename=next_master_page_style)
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
def create_header_footer(master_page, page_layout, header_or_footer, odd_or_even, nesting_level=0):
    header_footer = None
    header_footer_properties_attributes = {'margin': '0in', 'padding': '0in', 'dynamicspacing': False}
    mp_name = master_page.getAttribute('name')
    pl_name = page_layout.getAttribute('name')
    if header_or_footer == 'header':
        if odd_or_even == 'odd':
            header_footer = Header()
        elif odd_or_even == 'even':
            header_footer = HeaderLeft()
        elif odd_or_even == 'first':
            header_footer = Header()

        if header_footer:
            # TODO: the height should come from actual header content height
            header_style = style.HeaderStyle()
            header_style.addElement(style.HeaderFooterProperties(attributes=header_footer_properties_attributes))
            page_layout.addElement(header_style)
            master_page.addElement(header_footer)

    elif header_or_footer == 'footer':
        if odd_or_even == 'odd':
            header_footer = Footer()
        elif odd_or_even == 'even':
            header_footer = FooterLeft()
        elif odd_or_even == 'first':
            header_footer = Footer()

        if header_footer:
            # TODO: the height should come from actual header content height
            footer_style = style.FooterStyle()
            footer_style.addElement(style.HeaderFooterProperties(attributes=header_footer_properties_attributes))
            page_layout.addElement(footer_style)
            master_page.addElement(header_footer)

    # print(f"{header_or_footer}:{odd_or_even} going under {mp_name}:{pl_name}")

    return header_footer


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# various utility functions

''' process line-breaks
'''
def process_line_breaks(text, keep_line_breaks, nesting_level=0):
    if keep_line_breaks:
        new_text = text.replace('\n', '<text:line-break/>')
        return new_text

    else:
        return text.replace('\n', ' ')


''' given pixel size, calculate the row height in inches
    a reasonable approximation is what gsheet says 21 pixels, renders well as 12 pixel (assuming our normal text is 10-11 in size)
'''
def row_height_in_inches(pixel_size, nesting_level=0):
    return float((pixel_size) / 96)


''' get a random string
'''
def random_string(length=12, nesting_level=0):
    letters = string.ascii_uppercase
    return ''.join(random.choice(letters) for i in range(length))


''' fit width/height into a given width/height maintaining aspect ratio
'''
def fit_width_height(fit_within_width, fit_within_height, width_to_fit, height_to_fit, nesting_level=0):
	scale = min(
        fit_within_width / width_to_fit,
        fit_within_height / height_to_fit
    )

	return (
        width_to_fit * scale,
        height_to_fit * scale,
    )


''' strip LaTeX math mode delimeter ($)
'''
def strip_math_mode_delimeters(latex_content, nesting_level=0):
    # strip SPACES
    stripped = latex_content.strip()

    # strip $
    stripped = stripped.strip('$')

    # TODO: strip \( and \)

    return stripped


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ODT specific utility functions

''' registers a font in the document
'''
def register_font(odt, font_name, font_spec, nesting_level=0):
    font_family = font_spec
    trace(f"registering font [{font_name}] family [{font_spec}]", nesting_level=nesting_level+1)
    the_font = FontFace(name=font_name, fontfamily=font_family)
    odt.fontfacedecls.addElement(the_font)


''' return the style if exists
'''
def get_style_by_name(odt, style_name, nesting_level=0):
    # Check common styles
    for style in odt.styles.getElementsByType(Style):
        if style.getAttribute("name") == style_name:
            return style
        
    # Check automatic (local) styles
    for style in odt.automaticstyles.getElementsByType(Style):
        if style.getAttribute("name") == style_name:
            return style
    
    return None


''' apply a custom style to something
'''
def apply_custom_style(style, custom_properties, nesting_level=0):
    for p_type, props in PTOPERTY_TYPES.items():
        props_list_by_type = style.getElementsByType(PTOPERTY_TYPES[p_type])
        if props_list_by_type is None or len(props_list_by_type) == 0:
            klass = PTOPERTY_TYPES[p_type]
            obj = klass(attributes=custom_properties[p_type])
            style.addElement(obj)
        else:
            props_by_type = props_list_by_type[0]       
            if p_type in custom_properties:
                for attr, value in custom_properties[p_type].items():
                    trace(f"setting '{attr}' to [{value}]", nesting_level=nesting_level+1)
                    props_by_type.setAttribute(attr, value)



''' update a style from a given spec
'''
def update_style(odt, style_key, style_spec, custom_styles, nesting_level=0):
    custom_properties = parse_style_properties(style_spec=style_spec, nesting_level=nesting_level+1)

    style_name = style_spec.get('name', None)
    if style_name is None:
        trace(f"custom style [{style_key}] added to style cache")
        custom_styles[style_key] = custom_properties
        return
    
    else:
        style = get_style_by_name(odt=odt, style_name=style_name)
        if style is None:
            error(f"style [{style_name}] not found .. ")
            return

        # style exists, update with spec
        trace(f"overriding style [{style_name}]", nesting_level=nesting_level+1)
        for p_type, props in PTOPERTY_TYPES.items():
            props_by_type = style.getElementsByType(PTOPERTY_TYPES[p_type])[0]
            if p_type in custom_properties:
                for attr, value in custom_properties[p_type].items():
                    trace(f"setting '{attr}' to [{value}]", nesting_level=nesting_level+2)
                    props_by_type.setAttribute(attr, value)


''' parse style properties from yml to odt
'''
def parse_style_properties(style_spec, nesting_level=0):
    custom_properties = {}
    for p_type, props_dict in style_spec.items():
        if isinstance(props_dict, dict):
            new_props = {}
            for key, value in props_dict.items():
                if value and value != '':
                    new_key = key.split('.')[-1]
                    new_props[new_key] = value

            custom_properties[p_type] = new_props

    return custom_properties


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# various utility data

PTOPERTY_TYPES = {'text-properties': TextProperties, 'paragraph-properties': ParagraphProperties}

# seperation (in inches) between two ODT table columns
COLSEP = (0/72)

# seperation (in inches) between two ODT table rows
ROWSEP = (0/72)

# FACTOR by which to divide gsheet border width to get a reasonable ODT border width
ODT_BORDER_WIDTH_FACTOR = 4

# gsheet border style to ODT border style map
GSHEET_ODT_BORDER_MAPPING = {
    'DOTTED': 'dotted',
    'DASHED': 'dash',
    'SOLID': 'solid'
}


# ODT style name to outline level map
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

# outline level to ODT style name map
LEVEL_TO_HEADING = [
    'Title',
    'Heading_20_1',
    'Heading_20_2',
    'Heading_20_3',
    'Heading_20_4',
    'Heading_20_5',
    'Heading_20_6',
    'Heading_20_7',
    'Heading_20_8',
    'Heading_20_9',
    'Heading_20_10',
]


# gsheet text vertical alignment to ODT text vertical alignment map
TEXT_VALIGN_MAP = {'TOP': 'top', 'MIDDLE': 'middle', 'BOTTOM': 'bottom'}

# gsheet text horizontal alignment to ODT text horizontal alignment map
TEXT_HALIGN_MAP = {'LEFT': 'left', 'CENTER': 'center', 'RIGHT': 'right', 'JUSTIFY': 'justify'}

# gsheet image alignment to ODT image alignment map
IMAGE_POSITION = {'center': 'center', 'middle': 'center', 'left': 'left', 'right': 'right', 'top': 'top', 'bottom': 'bottom'}

# gsheet wrap strategy to ODT iwrap strategy map
WRAP_STRATEGY_MAP = {'OVERFLOW': 'no-wrap', 'CLIP': 'no-wrap', 'WRAP': 'wrap'}

# 0-based gsheet column number to column letter map
COLUMNS = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
            'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
            'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ']

PDF_PAGE_HEIGHT_OFFSET = 0.0

STYLE_PROPERTY_MAP = {
    "text-properties": 
    [
        "backgroundcolor",
        "color",
        "condition",
        {
            "country": [
                "country",
                "countryasian",
                "countrycomplex",
            ]
        },
        "display",
        {
            "font":
            [
                {
                    "charset": [
                        "fontcharset",
                        "fontcharsetasian",
                        "fontcharsetcomplex",
                    ],
                },   
                {
                    "family": [
                        "fontfamily",
                        "fontfamilyasian",
                        "fontfamilycomplex",
                        "fontfamilygeneric",
                        "fontfamilygenericasian",
                        "fontfamilygenericcomplex",
                    ],
                },
                {
                    "name": [
                        "fontname",
                        "fontnameasian",
                        "fontnamecomplex",
                    ],
                },   
                {
                    "pitch": [
                        "fontpitch",
                        "fontpitchasian",
                        "fontpitchcomplex",
                    ],
                },   
                "fontrelief",
                {
                    "size": [
                        "fontsize",
                        "fontsizeasian",
                        "fontsizecomplex",
                        "fontsizerel",
                        "fontsizerelasian",
                        "fontsizerelcomplex",
                    ],
                },   
                {
                    "style": [
                        "fontstyle",
                        "fontstyleasian",
                        "fontstylecomplex",
                        "fontstylename",
                        "fontstylenameasian",
                        "fontstylenamecomplex",
                    ],
                },   
                "fontvariant",
                {
                    "weight": [
                        "fontweight",
                        "fontweightasian",
                        "fontweightcomplex",
                    ],
                },
            ],
        },  
        "hyphenate",
        "hyphenationpushcharcount",
        "hyphenationremaincharcount",
        {
            "language": [
                "language",
                "languageasian",
                "languagecomplex",
            ],
        },   
        "letterkerning",
        "letterspacing",
        "scripttype",
        {
            "text": [
                "textblinking",
                "textcombine",
                "textcombineendchar",
                "textcombinestartchar",
                "textemphasize",
                "textlinethroughcolor",
                "textlinethroughmode",
                "textlinethroughstyle",
                "textlinethroughtext",
                "textlinethroughtextstyle",
                "textlinethroughtype",
                "textlinethroughwidth",
                "textoutline",
                "textposition",
                "textrotationangle",
                "textrotationscale",
                "textscale",
                "textshadow",
                "texttransform",
                "textunderlinecolor",
                "textunderlinemode",
                "textunderlinestyle",
                "textunderlinetype",
                "textunderlinewidth",
            ],
        },   
        "usewindowfontcolor",
    ],
    "paragraph-properties": [
        "autotextindent",
        "backgroundcolor",
        "backgroundtransparency",
        {
            "border": [
                "borderbottom",
                "borderleft",
                "borderlinewidth",
                "borderlinewidthbottom",
                "borderlinewidthleft",
                "borderlinewidthright",
                "borderlinewidthtop",
                "borderright",
                "bordertop",
            ],
        },
        "breakafter",
        "breakbefore",
        "fontindependentlinespacing",
        "hyphenationkeep",
        "hyphenationladdercount",
        "justifysingleword",
        "keeptogether",
        "keepwithnext",
        "linebreak",
        "lineheight",
        "lineheightatleast",
        "linenumber",
        "linespacing",
        {
            "margin": [
                "marginbottom",
                "marginleft",
                "marginright",
                "margintop",
            ],
        },
        "numberlines",
        "orphans",
        {
            "padding": [
                "paddingbottom",
                "paddingleft",
                "paddingright",
                "paddingtop",
            ],
        },
        "pagenumber",
        "punctuationwrap",
        "registertrue",
        "shadow",
        "snaptolayoutgrid",
        "tabstopdistance",
        "textalign",
        "textalignlast",
        "textautospace",
        "textindent",
        "verticalalign",
        "widows",
        "writingmode",
        "writingmodeautomatic",
    ],
}