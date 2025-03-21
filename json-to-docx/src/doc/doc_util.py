#!/usr/bin/env python3
'''
various utilities for formatting a docx
'''

import re
import sys
import lxml
import random
import string
import importlib

from pprint import pprint
from pathlib import Path
from copy import deepcopy

from lxml import etree

from docx import Document
from docx import section, document, table
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

from docx.shared import Pt, Cm, Inches, RGBColor, Emu

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_BREAK
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.enum.dml import MSO_THEME_COLOR_INDEX

import latex2mathml.converter

if sys.platform in ['win32', 'darwin']:
	import win32com.client as client

from helper.logger import *


''' process a list of section_data and generate docx code
'''
def section_list_to_doc(section_list, config):
	first_section = True
	for section in section_list:
		section_meta = section['section-meta']
		section_prop = section['section-prop']

		if section_prop['label'] != '':
			info(f"writing : {section_prop['label'].strip()} {section_prop['heading'].strip()}", nesting_level=section_meta['nesting-level'])
		else:
			info(f"writing : {section_prop['heading'].strip()}", nesting_level=section_meta['nesting-level'])

		module = importlib.import_module("doc.doc_api")
		func = getattr(module, f"process_{section_prop['content-type']}")
		func(section, config)


# --------------------------------------------------------------------------------------------------------------------------------------------
# pictures, background image

''' insert image into a container
'''
def insert_image(container, picture_path, width, height, bookmark=None):
	if is_table_cell(container):
		run = container.paragraphs[0].add_run()

		# bookmark
		if bookmark and bookmark != '':
			add_bookmark(paragraph=container.paragraphs[0], bookmark_name=bookmark, bookmark_text='')

		run.add_picture(picture_path, height=Inches(height), width=Inches(width))
		return container
	else:
		paragraph = container.add_paragraph()

		# bookmark
		if bookmark and bookmark != '':
			add_bookmark(paragraph=paragraph, bookmark_name=bookmark, bookmark_text='')

		run = paragraph.add_run()
		run.add_picture(picture_path, height=Inches(height), width=Inches(width))
		return paragraph


# --------------------------------------------------------------------------------------------------------------------------------------------
# table and table-cell

''' create a Table
'''
def create_table(container, num_rows, num_cols, container_width=None):
	tbl = None

	# debug(f"... creating ({num_rows} x {num_cols} : width = {container_width}) table for {type(container)}")
	if type(container) is section._Header or type(container) is section._Footer:
		# if the conrainer is a Header/Footer
		tbl = container.add_table(num_rows, num_cols, container_width)
	elif type(container) is table._Cell:
		# if the conrainer is a Cell
		tbl = container.add_table(num_rows, num_cols)
	elif type(container) is document.Document:
		# if the conrainer is a Document
		tbl = container.add_table(num_rows, num_cols)

	tbl.style = 'PlainTable'
	tbl.autofit = False

	return tbl


''' set repeat table row on every new page
'''
def set_repeat_table_header(row):
	tr = row._tr
	trPr = tr.get_or_add_trPr()
	tblHeader = OxmlElement('w:tblHeader')
	tblHeader.set(qn('w:val'), "true")
	trPr.append(tblHeader)
	return row


''' format container (paragraph or cell)
	border, bgcolor, padding, valign, halign
'''
def format_container(container, attributes, it_is_a_table_cell):
	# borders
	if it_is_a_table_cell:
		if 'padding' in attributes:
			set_cell_padding(container, padding=attributes['padding'])

		if 'borders' in attributes:
			set_cell_border(container, borders=attributes['borders'])

		if 'backgroundcolor' in attributes:
			set_cell_bgcolor(container, color=attributes['backgroundcolor'])

		if 'verticalalign' in attributes:
			container.vertical_alignment = attributes['verticalalign']

		if 'textalign' in attributes:
			container.paragraphs[0].alignment = attributes['textalign']

		# text rotation, only for table._cell
		if 'angle' in attributes:
			rotate_text(cell=container, direction=attributes['angle'])

	else:
		if 'borders' in attributes:
			set_paragraph_border(container, borders=attributes['borders'])

		if 'backgroundcolor' in attributes:
			set_paragraph_bgcolor(container, color=attributes['backgroundcolor'])

		if 'textalign' in attributes:
			container.alignment = attributes['textalign']


''' set table-cell borders
'''
def set_cell_border(cell: table._Cell, borders):
	tc = cell._tc
	tcPr = tc.get_or_add_tcPr()

	# check for tag existnace, if none found, then create one
	tcBorders = tcPr.first_child_found_in("w:tcBorders")
	if tcBorders is None:
		tcBorders = OxmlElement('w:tcBorders')
		tcPr.append(tcBorders)

	# list over all available tags
	for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
		edge_data = borders.get(edge)
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


''' set paragraph borders
'''
def set_paragraph_border(paragraph, borders):
	pPr = paragraph._p.get_or_add_pPr()

	# check for tag existnace, if none found, then create one
	pBorders = pPr.first_child_found_in("w:pBorders")
	if pBorders is None:
		pBorders = OxmlElement('w:pBorders')
		pPr.append(pBorders)

	# list over all available tags
	for edge in ('top', 'start', 'bottom', 'end'):
		edge_data = borders.get(edge)
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


''' set table-cell bgcolor
'''
def set_cell_bgcolor(cell: table._Cell, color):
	shading_elm_1 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color))
	cell._tc.get_or_add_tcPr().append(shading_elm_1)


''' set paragraph bgcolor
'''
def set_paragraph_bgcolor(paragraph, color):
	shading_elm_1 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color))
	paragraph._p.get_or_add_pPr().append(shading_elm_1)


''' set table-cell borders
'''
def set_cell_padding(cell: table._Cell, padding):
	tc = cell._tc
	tcPr = tc.get_or_add_tcPr()
	tcMar = OxmlElement('w:tcMar')

	for m in ["top", "start", "bottom", "end", ]:
		if m in padding:
			node = OxmlElement("w:{}".format(m))
			node.set(qn('w:w'), str(padding.get(m)))
			node.set(qn('w:type'), 'dxa')
			tcMar.append(node)

	tcPr.append(tcMar)


# --------------------------------------------------------------------------------------------------------------------------------------------
# paragraphs and texts

''' page number with/without page count
'''
def create_page_number(container, text_attributes=None, page_numbering='long', separator=' of '):
	paragraph = create_paragraph(container=container)
	run = paragraph.add_run()
	set_text_style(run, text_attributes)

	# page
	fldPage = OxmlElement('w:t')
	fldPage.set(qn('xml:space'), 'preserve')
	fldPage.text = 'Page '

	fldCharSeparate0 = OxmlElement('w:fldChar')
	fldCharSeparate0.set(qn('w:fldCharType'), 'separate')

	# create a new element and set attributes
	fldCharBegin1 = OxmlElement('w:fldChar')
	fldCharBegin1.set(qn('w:fldCharType'), 'begin')

	# actual page number
	instrText1 = OxmlElement('w:instrText')
	instrText1.set(qn('xml:space'), 'preserve')
	instrText1.text = 'PAGE \* MERGEFORMAT'

	fldCharSeparate1 = OxmlElement('w:fldChar')
	fldCharSeparate1.set(qn('w:fldCharType'), 'separate')

	fldCharEnd1 = OxmlElement('w:fldChar')
	fldCharEnd1.set(qn('w:fldCharType'), 'end')

	r_element = run._r
	r_element.append(fldPage)
	r_element.append(fldCharSeparate0)
	r_element.append(fldCharBegin1)
	r_element.append(instrText1)
	r_element.append(fldCharSeparate1)
	r_element.append(fldCharEnd1)

	if page_numbering == 'long':
		# number of pages
		fldCharOf = OxmlElement('w:t')
		fldCharOf.set(qn('xml:space'), 'preserve')
		fldCharOf.text = separator

		# create a new element and set attributes
		fldCharBegin2 = OxmlElement('w:fldChar')
		fldCharBegin2.set(qn('w:fldCharType'), 'begin')

		# numpages
		instrText2 = OxmlElement('w:instrText')
		instrText2.set(qn('xml:space'), 'preserve')
		instrText2.text = 'NUMPAGES \* MERGEFORMAT'

		fldCharSeparate2 = OxmlElement('w:fldChar')
		fldCharSeparate2.set(qn('w:fldCharType'), 'separate')

		fldCharEnd2 = OxmlElement('w:fldChar')
		fldCharEnd2.set(qn('w:fldCharType'), 'end')

		r_element.append(fldCharOf)
		r_element.append(fldCharBegin2)
		r_element.append(instrText2)
		r_element.append(fldCharSeparate2)
		r_element.append(fldCharEnd2)

	return paragraph


''' write a paragraph in a given style
'''
def create_paragraph(container, text_content=None, run_list=None, paragraph_attributes=None, text_attributes=None, outline_level=0, footnote_list={}, bookmark=None, directives=True):
	# create or get the paragraph
	if type(container) is section._Header or type(container) is section._Footer:
		# if the container is a Header/Footer
		paragraph = container.add_paragraph()

	elif type(container) is table._Cell:
		# if the conrainer is a Cell, the Cell already has an empty paragraph
		paragraph = container.paragraphs[0]

	elif type(container) is document.Document:
		# if the conrainer is a Document
		paragraph = container.add_paragraph()

	else:
		# if the conrainer is anything else
		paragraph = container.add_paragraph()


	# TODO: process inline blocks like footnotes, latex inside the content
	if run_list is not None:
		# run lists are a series of runs inside the paragraph
		for text_run in run_list:
			process_inline_blocks(paragraph=paragraph, text_content=text_run['text'], text_attributes=text_run['text-attributes'], footnote_list=footnote_list)

	elif text_content is not None:
		process_inline_blocks(paragraph=paragraph, text_content=text_content, text_attributes=text_attributes, footnote_list=footnote_list)


	# bookmark
	if bookmark and bookmark != '':
		add_bookmark(paragraph=paragraph, bookmark_name=bookmark, bookmark_text='')


	# apply the style if any
	if paragraph_attributes:
		# apply paragraph attrubutes
		if 'stylename' in paragraph_attributes:
			style_name = paragraph_attributes['stylename']
			paragraph.style = style_name

		pf = paragraph.paragraph_format

		# process new-page
		if 'breakbefore' in paragraph_attributes:
			pf.page_break_before = True

		# process keep-with-next
		if 'keepwithnext' in paragraph_attributes:
			pf.keep_with_next = True

		# process keep-with-previous
		if 'keepwithprevious' in paragraph_attributes:
			# docx does not support keep_with_previous
			# pf.keep_with_previous = True
			pass

	return paragraph


''' remove a paragraph
'''
def delete_paragraph(paragraph):
	p = paragraph._element
	p.getparent().remove(p)
	# p._p = p._element = None


''' process inline blocks inside a text and add to a paragraph
'''
def process_inline_blocks(paragraph, text_content, text_attributes, footnote_list):
    # process FN{...} first, we get a list of block dicts
    inline_blocks = process_footnotes(
        text_content=text_content, footnote_list=footnote_list
    )

    # process LATEX{...} for each text item
    new_inline_blocks = []
    for inline_block in inline_blocks:
        # process only 'text'
        if "text" in inline_block:
            new_inline_blocks = new_inline_blocks + process_latex_blocks(
                inline_block["text"]
            )

        else:
            new_inline_blocks.append(inline_block)

    inline_blocks = new_inline_blocks

    # process PAGE{..} for each text item
    new_inline_blocks = []
    for inline_block in inline_blocks:
        # process only 'text'
        if "text" in inline_block:
            new_inline_blocks = new_inline_blocks + process_bookmark_page_blocks(
                inline_block["text"]
            )

        else:
            new_inline_blocks.append(inline_block)

    inline_blocks = new_inline_blocks

    # we are ready to prepare the content
    for inline_block in inline_blocks:
        if "text" in inline_block:
            run = paragraph.add_run(inline_block["text"])
            set_text_style(run=run, text_attributes=text_attributes)

        elif "fn" in inline_block:
            create_footnote(paragraph=paragraph, footnote_tuple=inline_block["fn"])

        elif "latex" in inline_block:
            run = paragraph.add_run()
            create_latex(container=run, latex_content=inline_block["latex"])
            set_text_style(run=run, text_attributes=text_attributes)

        elif "page-num" in inline_block:
            run = add_page_reference(paragraph=paragraph, bookmark_name='')
            set_text_style(run=run, text_attributes=text_attributes)

        elif "page-count" in inline_block:
            run = add_page_reference(
                paragraph=paragraph, bookmark_name='*'
            )
            set_text_style(run=run, text_attributes=text_attributes)

        elif "bookmark-page" in inline_block:
            # add_link(paragraph=paragraph, link_to=inline_block["page"].strip(), text=inline_block["page"].strip(), tool_tip=None)
            run = add_page_reference(paragraph=paragraph, bookmark_name=inline_block["bookmark-page"].strip())
            set_text_style(run=run, text_attributes=text_attributes)


''' process footnotes inside text
'''
def process_footnotes(text_content, footnote_list):
	# if text contains footnotes we make a list containing texts->footnote->text->footnote ......
	texts_and_footnotes = []

	# find out if there is any match with FN#key inside the text_content
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
def process_latex_blocks(text_content):
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
def process_bookmark_page_blocks(text_content):
	# find out if there is any match with PAGE{...} inside the text_content
	texts_and_bookmarks = []

	pattern = r'PAGE{[^}]*}'
	current_index = 0
	for match in re.finditer(pattern, text_content):
		bookmark_content = match.group()[5:-1].strip()

		# we have found a PAGE block, we add the preceding text and the PAGE block into the list
		bookmark_start_index, bookmark_end_index = match.span()[0], match.span()[1]
		if bookmark_start_index >= current_index:
			# there are preceding text before the PAGE
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


''' create a footnote
'''
def create_footnote(paragraph, footnote_tuple):
	paragraph.add_footnote(footnote_tuple[1])


''' add a latex/mathml run into a paragraph
'''
def create_latex(container, latex_content):
	if latex_content is not None:
		mathml_output = latex2mathml.converter.convert(strip_math_mode_delimeters(latex_content))

		# Convert MathML (MML) into Office MathML (OMML) using a XSLT stylesheet
		tree = etree.fromstring(mathml_output)
		xslt = etree.parse(mml2omml_stylesheet_path)

		transform = etree.XSLT(xslt)
		new_dom = transform(tree)

		container._element.append(new_dom.getroot())


''' tex/character style for text run
'''
def set_text_style(run, text_attributes):
	if text_attributes:
		if 'fontweight' in text_attributes:
			run.bold = True
		else:
			run.bold = False

		if 'fontstyle' in text_attributes:
			run.italic = True
		else:
			run.italic = False

		if 'textlinethroughstyle' in text_attributes:
			run.strike = True
		else:
			run.strike = False

		if 'textunderlinestyle' in text_attributes:
			run.underline = True
		else:
			run.underline = False

		run.font.name = text_attributes['fontname']
		run.font.size = Pt(text_attributes['fontsize'])

		fgcolor = text_attributes.get('color')
		run.font.color.rgb = RGBColor(fgcolor.red, fgcolor.green, fgcolor.blue)


''' create a bookmark
'''
def add_bookmark(paragraph, bookmark_name, bookmark_text=''):
    run = paragraph.add_run()
    tag = run._r
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), '0')
    start.set(qn('w:name'), bookmark_name)
    tag.append(start)

    text = OxmlElement('w:r')
    text.text = bookmark_text
    tag.append(text)

    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), '0')
    end.set(qn('w:name'), bookmark_name)
    tag.append(end)


''' add a PAGE refernce 
'''
def add_page_reference(paragraph, bookmark_name):
    run = paragraph.add_run()

	# create a new element and set attributes
    fldCharBegin = OxmlElement('w:fldChar')
    fldCharBegin.set(qn('w:fldCharType'), 'begin')

    # actual PAGE reference
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")

    if bookmark_name == '':
        instrText.text = 'PAGE \* MERGEFORMAT'

    elif bookmark_name == '*':
        instrText.text = 'NUMPAGES \* MERGEFORMAT'

    else:
        instrText.text = f"PAGEREF {bookmark_name} \\h"

    fldCharEnd = OxmlElement('w:fldChar')
    fldCharEnd.set(qn('w:fldCharType'), 'end')

    r_element = run._r
    r_element.append(fldCharBegin)
    r_element.append(instrText)
    r_element.append(fldCharEnd)

    # p_element = paragraph._p

    return run


''' add link for bookmarks
'''
def add_link(paragraph, link_to, text, tool_tip=None):
	# create hyperlink node
	hyperlink = OxmlElement('w:hyperlink')

	# set attribute for link to bookmark
	hyperlink.set(qn('w:anchor'), link_to,)

	if tool_tip is not None:
		# set attribute for link to bookmark
		hyperlink.set(qn('w:tooltip'), tool_tip,)

	new_run = OxmlElement('w:r')
	rPr = OxmlElement('w:rPr')
	new_run.append(rPr)
	new_run.text = text
	hyperlink.append(new_run)
	r = paragraph.add_run()
	r._r.append(hyperlink)
	r.font.name = "Calibri"
	r.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
	r.font.underline = True


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# indexes and pdf generation

''' update docx indexes by opening and closing the docx, rest is done by macros
'''
def update_indexes(docx_path):

	try:
		word = client.DispatchEx("Word.Application")
	except Exception as e:
		raise e

	try:
		doc = word.Documents.Open(docx_path)
		doc.Close()
	except Exception as e:
		raise e
	finally:
		word.Quit()


''' set docx updateFields property true
'''
def set_updatefields_true(docx_path):
	namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
	doc = Document(docx_path)
	# add child to doc.settings element
	element_updatefields = lxml.etree.SubElement(
		doc.settings.element, f"{namespace}updateFields"
	)
	element_updatefields.set(f"{namespace}val", "true")
	doc.save(docx_path)


''' given an docx file generates pdf in the given directory
'''
def generate_pdf(infile, outdir):
	pdf_path = infile + '.pdf'
	try:
		word = client.DispatchEx("Word.Application")
	except Exception as e:
		raise e

	try:
		doc = word.Documents.Open(infile)
		try:
			doc.SaveAs(pdf_path, FileFormat = 17)

		except Exception as e:
			raise e

		finally:
			doc.Close()

	except Exception as e:
		raise e

	finally:
		word.Quit()


''' create table-of-contents
'''
def create_index(doc, index_type):
	paragraph = doc.add_paragraph()
	run = paragraph.add_run()

	# create a new element with attributes
	fldChar = OxmlElement('w:fldChar')
	fldChar.set(qn('w:fldCharType'), 'begin')

	instrText = OxmlElement('w:instrText')
	instrText.set(qn('xml:space'), 'preserve')

	if index_type == 'toc':
		instrText.text = 'TOC \\o "1-6" \\h \\z \\u'
	elif index_type == 'lof':
		instrText.text = 'TOC \\h \\z \\t "Figure" \\c'
	elif index_type == 'lot':
		instrText.text = 'TOC \\h \\z \\t "Table" \\c'

	fldChar2 = OxmlElement('w:fldChar')
	fldChar2.set(qn('w:fldCharType'), 'separate')
	fldChar3 = OxmlElement('w:t')
	fldChar3.text = "Right-click to update Index."
	fldChar2.append(fldChar3)

	fldChar4 = OxmlElement('w:fldChar')
	fldChar4.set(qn('w:fldCharType'), 'end')

	r_element = run._r
	r_element.append(fldChar)
	r_element.append(instrText)
	r_element.append(fldChar2)
	r_element.append(fldChar4)


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# document-section, page-layout, header-footer

''' add or update a document section
'''
def add_or_update_document_section(doc, page_spec, margin_spec, orientation, different_firstpage, section_break, page_break, first_section, different_odd_even_pages, background_image_path):
	#  if it is a section break, we isnert a new section
	if section_break:
		new_section = True
		docx_section = doc.add_section(WD_SECTION.NEW_PAGE)

	else:
		# we are continuing the last section
		if first_section:
			new_section = True
		else:
			new_section = False

		docx_section = doc.sections[-1]

		#  we may have a page break
		if page_break:
			doc.add_page_break()


	docx_section.first_page_header.is_linked_to_previous = False
	docx_section.header.is_linked_to_previous = False
	docx_section.even_page_header.is_linked_to_previous = False

	docx_section.first_page_footer.is_linked_to_previous = False
	docx_section.footer.is_linked_to_previous = False
	docx_section.even_page_footer.is_linked_to_previous = False

	if orientation == 'landscape':
		docx_section.orient = WD_ORIENT.LANDSCAPE
		docx_section.page_width = Inches(page_spec['height'])
		docx_section.page_height = Inches(page_spec['width'])
	else:
		docx_section.orient = WD_ORIENT.PORTRAIT
		docx_section.page_width = Inches(page_spec['width'])
		docx_section.page_height = Inches(page_spec['height'])

	docx_section.left_margin = Inches(margin_spec['left'])
	docx_section.right_margin = Inches(margin_spec['right'])
	docx_section.top_margin = Inches(margin_spec['top'])
	docx_section.bottom_margin = Inches(margin_spec['bottom'])

	docx_section.gutter = Inches(margin_spec['gutter'])

	docx_section.header_distance = Inches(margin_spec['distance']['header'])
	docx_section.footer_distance = Inches(margin_spec['distance']['footer'])

	docx_section.different_first_page_header_footer = different_firstpage
	doc.settings.odd_and_even_pages_header_footer = different_odd_even_pages

	# get the actual width
	actual_width = docx_section.page_width.inches - docx_section.left_margin.inches - docx_section.right_margin.inches - docx_section.gutter.inches

	# TODO: background-image
	if background_image_path != '':
		create_page_background(doc=doc, background_image_path=background_image_path, page_width_inches=docx_section.page_width.inches, page_height_inches=docx_section.page_height.inches)

	return docx_section, new_section


''' create page background
	<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
		relativeHeight="251658240" behindDoc="1" locked="0" layoutInCell="1"
		allowOverlap="1">
		<wp:simplePos x="0" y="0" />
		<wp:positionH relativeFrom="page">
			<wp:posOffset>0</wp:posOffset>
		</wp:positionH>
		<wp:positionV relativeFrom="page">
			<wp:posOffset>0</wp:posOffset>
		</wp:positionV>
		<wp:extent cx="7562000" cy="10689336" />
		<wp:effectExtent l="0" t="0" r="0" b="0" />
		<wp:wrapNone />
		<wp:docPr id="1" name="Picture 1" />
		<wp:cNvGraphicFramePr>
			<a:graphicFrameLocks
				xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
		</wp:cNvGraphicFramePr>
		<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<a:graphicData
				uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
				<pic:pic
					xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
					<pic:nvPicPr>
						<pic:cNvPr id="0" name="toll-booth.jpg" />
						<pic:cNvPicPr />
					</pic:nvPicPr>
					<pic:blipFill>
						<a:blip r:embed="rId8">
							<a:extLst>
								<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
									<a14:useLocalDpi
										xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
										val="0" />
								</a:ext>
							</a:extLst>
						</a:blip>
						<a:stretch>
							<a:fillRect />
						</a:stretch>
					</pic:blipFill>
					<pic:spPr>
						<a:xfrm>
							<a:off x="0" y="0" />
							<a:ext cx="7562000" cy="10692000" />
						</a:xfrm>
						<a:prstGeom prst="rect">
							<a:avLst />
						</a:prstGeom>
					</pic:spPr>
				</pic:pic>
			</a:graphicData>
		</a:graphic>
		<wp14:sizeRelH relativeFrom="margin">
			<wp14:pctWidth>0</wp14:pctWidth>
		</wp14:sizeRelH>
		<wp14:sizeRelV relativeFrom="margin">
			<wp14:pctHeight>0</wp14:pctHeight>
		</wp14:sizeRelV>
	</wp:anchor>
'''
def create_page_background(doc, background_image_path, page_width_inches, page_height_inches):
	drawing_xml = '''
	<w:drawing>
		<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
			<wp:simplePos x="0" y="0" />
			<wp:positionH relativeFrom="page">
				<wp:posOffset>0</wp:posOffset>
			</wp:positionH>
			<wp:positionV relativeFrom="page">
				<wp:posOffset>0</wp:posOffset>
			</wp:positionV>
			<wp:extent cx="{cx}" cy="{cy}" />
			<wp:effectExtent l="0" t="0" r="0" b="0" />
			<wp:wrapNone />
			<wp:docPr id="{doc_id}" name="{doc_name}" />
			<wp:cNvGraphicFramePr>
				<a:graphicFrameLocks
					xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
			</wp:cNvGraphicFramePr>
			<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
				<a:graphicData
					uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
					<pic:pic
						xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
						<pic:nvPicPr>
							<pic:cNvPr id="{image_id}" name="{image_name}" />
							<pic:cNvPicPr />
						</pic:nvPicPr>
						<pic:blipFill>
							<a:blip r:embed="{rid}">
							</a:blip>
							<a:stretch>
								<a:fillRect />
							</a:stretch>
						</pic:blipFill>
						<pic:spPr>
							<a:xfrm>
								<a:off x="0" y="0" />
								<a:ext cx="{cx}" cy="{cy}" />
							</a:xfrm>
							<a:prstGeom prst="rect">
								<a:avLst />
							</a:prstGeom>
						</pic:spPr>
					</pic:pic>
				</a:graphicData>
			</a:graphic>
			<wp14:sizeRelH relativeFrom="margin">
				<wp14:pctWidth>0</wp14:pctWidth>
			</wp14:sizeRelH>
			<wp14:sizeRelV relativeFrom="margin">
				<wp14:pctHeight>0</wp14:pctHeight>
			</wp14:sizeRelV>
		</wp:anchor>
	</w:drawing>
	'''

	# get the current/last paragraph
	first_para = doc.paragraphs[-1]
	first_run = first_para.add_run()

	# add a new paragraph paragraph
	new_para = doc.add_paragraph()

	# create a run
	new_run = new_para.add_run()

	# image
	# shape = r.add_picture(image_path_or_stream=background_image_path, width=docx_section.page_width.inches, height=docx_section.page_height.inches)
	shape = new_run.add_picture(image_path_or_stream=background_image_path)
	current_drawing_element = new_run._r.xpath('//w:drawing')[0]


	# tweak the generated inline image
	parser = etree.XMLParser(recover=True)

	cx = int(EMU_PER_INCH * (page_width_inches - 0.0))
	cy = int(EMU_PER_INCH * (page_height_inches - 0.0))

	docPr = new_run._r.xpath('//wp:docPr')[0]
	doc_id = docPr.get('id')
	doc_name = docPr.get('name')

	cNvPr = new_run._r.xpath('//pic:cNvPr')[0]
	image_id = cNvPr.get('id')
	image_name = cNvPr.get('name')

	blip = new_run._r.xpath('//a:blip')[0]
	rid = blip.xpath('./@r:embed')[0]

	new_drawing_element = etree.XML(drawing_xml.format(cx=cx, cy=cy, doc_id=doc_id, doc_name=doc_name, image_id=image_id, image_name=image_name, rid=rid), parser)


	# remove the new-para
	delete_paragraph(new_para)

	# put the new drawing into first-run
	first_run._r.append(new_drawing_element)


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# various utility functions

''' whether the container is a table-cell or not
'''
def is_table_cell(container):
	# if container is n instance of table-cell
	if type(container) is table._Cell:
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
	HEIGHT_OFFSET = 0.3

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


'''
'''
def strip_math_mode_delimeters(latex_content):
	# strip SPACES
	stripped = latex_content.strip()

	# strip $
	stripped = stripped.strip('$')

	# TODO: strip \( and \)

	return stripped


# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# various utility data

EMU_PER_INCH = 914500

mml2omml_stylesheet_path = '../conf/MML2OMML_15.XSL'
# mml2omml_stylesheet_path = '../conf/MML2OMML_16.XSL'


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


TEXT_VALIGN_MAP = {
	'TOP': WD_CELL_VERTICAL_ALIGNMENT.TOP,
	'MIDDLE': WD_CELL_VERTICAL_ALIGNMENT.CENTER,
	'BOTTOM': WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
}


TEXT_HALIGN_MAP = {
	'LEFT': WD_ALIGN_PARAGRAPH.LEFT,
	'CENTER': WD_ALIGN_PARAGRAPH.CENTER,
	'RIGHT': WD_ALIGN_PARAGRAPH.RIGHT,
	'JUSTIFY': WD_ALIGN_PARAGRAPH.JUSTIFY
}


WRAP_STRATEGY_MAP = {'OVERFLOW': 'no-wrap', 'CLIP': 'no-wrap', 'WRAP': 'wrap'}


COLUMNS = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
			'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
			'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ']


GSHEET_OXML_BORDER_MAPPING = {
	'DOTTED': 'dotted',
	'DASHED': 'dashed',
	'SOLID': 'single',
	'SOLID_MEDIUM': 'thick',
	'SOLID_THICK': 'triple',
	'DOUBLE': 'double',
	'NONE': 'none'
}


def rotate_text(cell: table._Cell, direction: str):
	# direction: tbRl -- top to bottom, btLr -- bottom to top
	assert direction in ("tbRl", "btLr")
	tc = cell._tc
	tcPr = tc.get_or_add_tcPr()
	textDirection = OxmlElement('w:textDirection')
	textDirection.set(qn('w:val'), direction)  # btLr tbRl
	tcPr.append(textDirection)


def copy_cell_border(from_cell: table._Cell, to_cell: table._Cell):
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
