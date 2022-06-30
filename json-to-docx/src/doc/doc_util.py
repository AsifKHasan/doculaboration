#!/usr/bin/env python3
'''
various utilities for formatting a docx
'''

import sys
import lxml
import random
import string
from pathlib import Path
from copy import deepcopy

from docx import Document
from docx import section, document, table
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

from docx.shared import Pt, Cm, Inches, RGBColor, Emu

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_BREAK
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION, WD_ORIENT

if sys.platform == 'win32':
	import win32com.client as client

from helper.logger import *


# --------------------------------------------------------------------------------------------------------------------------------------------
# pictures, background image

''' insert image into a container
'''
def insert_image(container, picture_path, width, height):
	if is_table_cell(container):
		run = container.paragraphs[0].add_run()
		run.add_picture(picture_path, height=Inches(height), width=Inches(width))
		return container
	else:
		paragraph = container.add_paragraph()
		run = paragraph.add_run()
		run.add_picture(picture_path, height=Inches(height), width=Inches(width))
		return paragraph



# --------------------------------------------------------------------------------------------------------------------------------------------
# table and table-cell

''' create a Table
'''
def create_table(container, num_rows, num_cols, container_width=None):
	tbl = None

	debug(f"... creating ({num_rows} x {num_cols}) table for {type(container)}")
	if type(container) is section._Header or type(container) is section._Footer:
		# if the conrainer is a Header/Footer
		tbl = container.add_table(num_rows, num_cols, container_width)
	elif type(container) is table._Cell:
		# if the conrainer is a Cell
		tbl = container.add_table(num_rows, num_cols)
	elif type(container) is document.Document:
		# if the conrainer is a Document
		tbl = container.add_table(num_rows, num_cols)

	return tbl


def format_container(container, attributes):
	# borders
	if 'borders' in attributes:
		if is_table_cell(container):
			print(f".. cell bordering")
			set_cell_border(container, borders=attributes['borders'])
		else:
			print(f".. paragraph bordering")
			set_paragraph_border(container, borders=attributes['borders'])



''' set table-cell borders
	top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
	bottom={"sz": 12, "color": "#00FF00", "val": "single"},
	start={"sz": 24, "val": "dashed", "shadow": "true"},
	end={"sz": 12, "val": "dashed"},
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
			print(f".... found cell edge : {edge}")
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
	top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
	bottom={"sz": 12, "color": "#00FF00", "val": "single"},
	start={"sz": 24, "val": "dashed", "shadow": "true"},
	end={"sz": 12, "val": "dashed"},
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
			print(f".... found paragraph edge : {edge}")
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



# --------------------------------------------------------------------------------------------------------------------------------------------
# paragraphs and texts

''' page number wit/without page count
'''
def create_page_number(paragraph, short=False, separator=' of '):
	run = paragraph.add_run()

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
	r_element.append(fldCharBegin1)
	r_element.append(instrText1)
	r_element.append(fldCharSeparate1)
	r_element.append(fldCharEnd1)

	if not short:
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

	p_element = paragraph._p


''' write a paragraph in a given style
'''
def create_paragraph(container, text_content=None, run_list=None, style_attributes=None, paragraph_attributes=None, text_attributes=None, outline_level=0):
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


	# apply the style if any
	if style_attributes and 'parentstylename' in style_attributes:
		style_name = style_attributes['parentstylename']
		paragraph.style = style_name


	# apply paragraph attrubutes
	if paragraph_attributes:
		pf = paragraph.paragraph_format
		# process new-page
		if 'breakbefore' in paragraph_attributes:
			pf.page_break_before = True

		# process keep-with-next
		if 'keepwithnext' in paragraph_attributes:
			pf.keep_with_next = True

		# process keep-with-previous
		if 'keepwithprevious' in paragraph_attributes:
			pf.keep_with_previous = True


	# run lists are a series of runs inside the paragraph
	if run_list is not None:
		for text_run in run_list:
			run = paragraph.add_run(text_run['text'])
			set_text_style(run, text_run['text-attributes'])

	else:
		run = paragraph.add_run(text_content)
		set_text_style(run, text_attributes)

	return paragraph


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








# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# indexes and pdf generation

''' update docx indexes by opening and closing the docx, rest is done by macros
'''
def update_indexes(docx_path):
	try:
		word = client.DispatchEx("Word.Application")
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
	pdf_path = infile.replace(".docx", r".pdf")
	try:
		word = client.DispatchEx("Word.Application")
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
	p_element = paragraph._p



# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# document-section, page-layout, header-footer

''' add or update a document section
'''
def add_or_update_document_section(doc, docx_specs, page_spec, margin_spec, orientation, different_firstpage, section_index=None):
	if not section_index is None:
		# we want to change section with index section_index
		section = doc.sections[section_index]
	else:
		# new section always starts with a page-break
		section = doc.add_section(WD_SECTION.NEW_PAGE)

	section.first_page_header.is_linked_to_previous = False
	section.header.is_linked_to_previous = False
	section.even_page_header.is_linked_to_previous = False

	section.first_page_footer.is_linked_to_previous = False
	section.footer.is_linked_to_previous = False
	section.even_page_footer.is_linked_to_previous = False

	if orientation == 'landscape':
		section.orient = WD_ORIENT.LANDSCAPE
		section.page_width = Inches(docx_specs['page-spec'][page_spec]['height'])
		section.page_height = Inches(docx_specs['page-spec'][page_spec]['width'])
	else:
		section.orient = WD_ORIENT.PORTRAIT
		section.page_width = Inches(docx_specs['page-spec'][page_spec]['width'])
		section.page_height = Inches(docx_specs['page-spec'][page_spec]['height'])

	section.left_margin = Inches(docx_specs['margin-spec'][margin_spec]['left'])
	section.right_margin = Inches(docx_specs['margin-spec'][margin_spec]['right'])
	section.top_margin = Inches(docx_specs['margin-spec'][margin_spec]['top'])
	section.bottom_margin = Inches(docx_specs['margin-spec'][margin_spec]['bottom'])

	section.gutter = Inches(docx_specs['margin-spec'][margin_spec]['gutter'])

	section.header_distance = Inches(docx_specs['margin-spec'][margin_spec]['distance']['header'])
	section.footer_distance = Inches(docx_specs['margin-spec'][margin_spec]['distance']['footer'])

	section.different_first_page_header_footer = different_firstpage

	# get the actual width
	actual_width = section.page_width.inches - section.left_margin.inches - section.right_margin.inches - section.gutter.inches

	return section



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


def set_cell_bgcolor(cell, color):
	shading_elm_1 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color))
	cell._tc.get_or_add_tcPr().append(shading_elm_1)


def set_paragraph_bgcolor(paragraph, color):
	shading_elm_1 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color))
	paragraph._p.get_or_add_pPr().append(shading_elm_1)


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


def set_repeat_table_header(row):
	''' set repeat table row on every new page
	'''
	tr = row._tr
	trPr = tr.get_or_add_trPr()
	tblHeader = OxmlElement('w:tblHeader')
	tblHeader.set(qn('w:val'), "true")
	trPr.append(tblHeader)
	return row
