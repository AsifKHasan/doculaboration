#!/usr/bin/env python3

'''
various utilities for rendering gsheet cell content into a docx, mostly for Formatter of type Table
'''

import json
import time
import pprint

from docx.shared import Pt, Cm, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_BREAK
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION, WD_ORIENT

from helper.logger import *
from helper.docx.docx_util import *

VALIGN = {'TOP': WD_CELL_VERTICAL_ALIGNMENT.TOP, 'MIDDLE': WD_CELL_VERTICAL_ALIGNMENT.CENTER, 'BOTTOM': WD_CELL_VERTICAL_ALIGNMENT.BOTTOM}
HALIGN = {'LEFT': WD_ALIGN_PARAGRAPH.LEFT, 'CENTER': WD_ALIGN_PARAGRAPH.CENTER, 'RIGHT': WD_ALIGN_PARAGRAPH.RIGHT, 'JUSTIFY': WD_ALIGN_PARAGRAPH.JUSTIFY}

def render_content_in_cell(doc, cell, cell_data, width, r, c, start_row, start_col, merge_data, column_widths, table_spacing):
    cell.width = Inches(width)
    paragraph = cell.paragraphs[0]

    # paragraph spacing
    if table_spacing == 'no-spacing':
        pf = paragraph.paragraph_format
        pf.line_spacing = 1.0
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

    # handle the notes first
    # if there is a note, see if it is a JSON, it may contain style, page-numering, new-page, keep-with-next directive etc.
    note_json = {}
    if 'note' in cell_data:
        try:
            note_json = json.loads(cell_data['note'])
        except json.JSONDecodeError:
            pass

    # process new-page
    if 'new-page' in note_json:
        # return the cell location so that the page break can be rendered later
        pf = paragraph.paragraph_format
        pf.page_break_before = True

    # process keep-with-next
    if 'keep-with-next' in note_json:
        # return the cell location so that the page break can be rendered later
        pf = paragraph.paragraph_format
        pf.keep_with_next = True

    # do some special processing if the cell_data is {}
    if cell_data == {} or 'effectiveFormat' not in cell_data:
        return

    text_format = cell_data['effectiveFormat']['textFormat']
    effective_format = cell_data['effectiveFormat']

    # alignments
    cell.vertical_alignment = VALIGN[effective_format['verticalAlignment']]
    if 'horizontalAlignment' in effective_format:
        paragraph.alignment = HALIGN[effective_format['horizontalAlignment']]

    # background color
    bgcolor = cell_data['effectiveFormat']['backgroundColor']
    if bgcolor != {}:
        red = int(bgcolor['red'] * 255) if 'red' in bgcolor else 0
        green = int(bgcolor['green'] * 255) if 'green' in bgcolor else 0
        blue = int(bgcolor['blue'] * 255) if 'blue' in bgcolor else 0
        set_cell_bgcolor(cell, RGBColor(red, green, blue))

    # text-rotation
    if 'textRotation' in effective_format:
        text_rotation = effective_format['textRotation']
        rotate_text(cell, 'btLr')

    # borders
    if 'borders' in cell_data['effectiveFormat']:
        borders = cell_data['effectiveFormat']['borders']
        set_cell_border(cell, top=ooxml_border_from_gsheet_border(borders, 'top'), bottom=ooxml_border_from_gsheet_border(borders, 'bottom'), start=ooxml_border_from_gsheet_border(borders, 'left'), end=ooxml_border_from_gsheet_border(borders, 'right'))

    # cell can be merged, so we need width after merge (in Inches)
    cell_width = merged_cell_width(r, c, start_row, start_col, merge_data, column_widths)

    # images
    if 'userEnteredValue' in cell_data:
        userEnteredValue = cell_data['userEnteredValue']
        if 'image' in userEnteredValue:
            image = userEnteredValue['image']
            run = paragraph.add_run()

            # even now the width may exceed actual cell width, we need to adjust for that
            # determine cell_width based on merge scenario
            dpi_x = 150 if image['dpi'][0] == 0 else image['dpi'][0]
            dpi_y = 150 if image['dpi'][1] == 0 else image['dpi'][1]
            image_width = image['width'] / dpi_x
            image_height = image['height'] / dpi_y
            if image_width > cell_width:
                adjust_ratio = (cell_width / image_width)
                # keep a padding of 0.1 inch
                image_width = cell_width - 0.2
                image_height = image_height * adjust_ratio

            run.add_picture(image['path'], height=Inches(image_height), width=Inches(image_width))

    # before rendering cell, see if it embeds another worksheet
    if 'contents' in cell_data:
        table = insert_content(cell_data['contents'], doc, cell_width, container=None, cell=cell)
        # polish_table(table)
        return

    # texts
    if 'formattedValue' not in cell_data:
        return

    text = cell_data['formattedValue']

    # process notes
    # note specifies style
    if 'style' in note_json:
        paragraph.add_run(text)
        paragraph.style = note_json['style']
        return

    # note specifies page numbering
    if 'page-number' in note_json:
        append_page_number_with_pages(paragraph)
        #append_page_number_only(paragraph)
        paragraph.style = note_json['page-number']
        return

    # finally cell content, add runs
    if 'textFormatRuns' in cell_data:
        text_runs = cell_data['textFormatRuns']
        # split the text into run-texts
        run_texts = []
        for i in range(len(text_runs) - 1, -1, -1):
            text_run = text_runs[i]
            if 'startIndex' in text_run:
                run_texts.insert(0, text[text_run['startIndex']:])
                text = text[:text_run['startIndex']]
            else:
                run_texts.insert(0, text)

        # now render runs
        for i in range(0, len(text_runs)):
            # get formatting
            format = text_runs[i]['format']

            run = paragraph.add_run(run_texts[i])
            set_character_style(run, {**text_format, **format})
    else:
        run = paragraph.add_run(text)
        set_character_style(run, text_format)


def render_content_in_doc(doc, cell_data):
    paragraph = doc.add_paragraph()

    # handle the notes first
    # if there is a note, see if it is a JSON, it may contain style, page-numering, new-page, keep-with-next directive etc.
    note_json = {}
    if 'note' in cell_data:
        try:
            note_json = json.loads(cell_data['note'])
        except json.JSONDecodeError:
            pass

    # process new-page
    if 'new-page' in note_json:
        # return the cell location so that the page break can be rendered later
        pf = paragraph.paragraph_format
        pf.page_break_before = True

    # process keep-with-next
    if 'keep-with-next' in note_json:
        # return the cell location so that the page break can be rendered later
        pf = paragraph.paragraph_format
        pf.keep_with_next = True

    # do some special processing if the cell_data is {}
    if cell_data == {} or 'effectiveFormat' not in cell_data:
        return

    text_format = cell_data['effectiveFormat']['textFormat']
    effective_format = cell_data['effectiveFormat']

    # alignments
    # cell.vertical_alignment = VALIGN[effective_format['verticalAlignment']]
    if 'horizontalAlignment' in effective_format:
        paragraph.alignment = HALIGN[effective_format['horizontalAlignment']]

    # borders
    if 'borders' in cell_data['effectiveFormat']:
        borders = cell_data['effectiveFormat']['borders']
        set_paragraph_border(paragraph, top=ooxml_border_from_gsheet_border(borders, 'top'), bottom=ooxml_border_from_gsheet_border(borders, 'bottom'), start=ooxml_border_from_gsheet_border(borders, 'left'), end=ooxml_border_from_gsheet_border(borders, 'right'))

    # background color
    bgcolor = cell_data['effectiveFormat']['backgroundColor']
    if bgcolor != {}:
        red = int(bgcolor['red'] * 255) if 'red' in bgcolor else 0
        green = int(bgcolor['green'] * 255) if 'green' in bgcolor else 0
        blue = int(bgcolor['blue'] * 255) if 'blue' in bgcolor else 0
        set_paragraph_bgcolor(paragraph, RGBColor(red, green, blue))

    # images
    if 'userEnteredValue' in cell_data:
        userEnteredValue = cell_data['userEnteredValue']
        if 'image' in userEnteredValue:
            image = userEnteredValue['image']
            run = paragraph.add_run()

            # even now the width may exceed actual cell width, we need to adjust for that
            # determine cell_width based on merge scenario
            dpi_x = 150 if image['dpi'][0] == 0 else image['dpi'][0]
            dpi_y = 150 if image['dpi'][1] == 0 else image['dpi'][1]
            image_width = image['width'] / dpi_x
            image_height = image['height'] / dpi_y
            if image_width > cell_width:
                adjust_ratio = (cell_width / image_width)
                # keep a padding of 0.1 inch
                image_width = cell_width - 0.2
                image_height = image_height * adjust_ratio

            run.add_picture(image['path'], height=Inches(image_height), width=Inches(image_width))

    # before rendering cell, see if it embeds another worksheet
    if 'contents' in cell_data:
        table = insert_content(cell_data['contents'], doc, cell_width, container=None, cell=cell)
        # polish_table(table)
        return

    # texts
    if 'formattedValue' not in cell_data:
        return

    text = cell_data['formattedValue']

    # process notes
    # note specifies style
    if 'style' in note_json:
        paragraph.add_run(text)
        paragraph.style = note_json['style']
        return

    # note specifies page numbering
    if 'page-number' in note_json:
        append_page_number_with_pages(paragraph)
        #append_page_number_only(paragraph)
        paragraph.style = note_json['page-number']
        return

    # finally cell content, add runs
    if 'textFormatRuns' in cell_data:
        text_runs = cell_data['textFormatRuns']
        # split the text into run-texts
        run_texts = []
        for i in range(len(text_runs) - 1, -1, -1):
            text_run = text_runs[i]
            if 'startIndex' in text_run:
                run_texts.insert(0, text[text_run['startIndex']:])
                text = text[:text_run['startIndex']]
            else:
                run_texts.insert(0, text)

        # now render runs
        for i in range(0, len(text_runs)):
            # get formatting
            format = text_runs[i]['format']

            run = paragraph.add_run(run_texts[i])
            set_character_style(run, {**text_format, **format})
    else:
        run = paragraph.add_run(text)
        set_character_style(run, text_format)


def insert_content(data, doc, container_width, container=None, cell=None, repeat_rows=0):
    if not container: debug('.. inserting contents')

    # we have a concept of in-cell content and out-of-cell content
    # in-cell content means the content will go inside an existing table cell (specified by a 'content' key with value anything but 'out-of-cell' or no 'content' key at all in 'notes')
    # out-of-cell content means the content will be inserted as a new table directly in the doc (specified by a 'content' key with value 'out-of-cell' in 'notes')

    # when there is an out-of-cell content, we need to put the content inside the document by getting out of any table cell if we are already inside any cell
    # this means we need to look into the first column of each row and identify if we have an out-of-cell content
    # if we have such a content, anything prior to this content will go in one table, the out-of-cell content will go into the document and subsequent cells will go into another table after the put-of-cell content
    # we basically need content segmentation/segrefation into parts

    start_row, start_col = data['sheets'][0]['data'][0]['startRow'], data['sheets'][0]['data'][0]['startColumn']
    worksheet_rows = data['sheets'][0]['properties']['gridProperties']['rowCount']
    worksheet_cols = data['sheets'][0]['properties']['gridProperties']['columnCount']

    row_data = data['sheets'][0]['data'][0]['rowData']
    out_of_cell_content_rows = []
    # we are looking for out-of-cell content
    for row_num in range(start_row + 1, worksheet_rows + 1):
        # we look only in the first column
        # debug('at row : {0}/{1}'.format(row_num, worksheet_rows))
        row_data_index = row_num - start_row - 1

        # get the first cell notes
        first_cell_data = row_data[row_data_index]['values'][0]
        first_cell_note_json = {}
        if 'note' in first_cell_data:
            try:
                first_cell_note_json = json.loads(first_cell_data['note'])
            except json.JSONDecodeError:
                pass

        # see the 'content' tag
        out_of_cell_content = False
        if 'content' in first_cell_note_json:
            if first_cell_note_json['content'] == 'out-of-cell':
                out_of_cell_content = True

        # if content is out of cell, mark the row
        if out_of_cell_content:
            out_of_cell_content_rows.append(row_num)

    # we have found all rows which are to be rendered out-of-cell that is directly into the doc, not inside any existing table cell
    # now we have to segment the rows into table segments and out-of-cell segments
    start_at_row = start_row + 1
    row_segments = []
    for row_num in out_of_cell_content_rows:
        if row_num > start_at_row:
            row_segments.append({'table': (start_at_row, row_num - 1)})

        row_segments.append({'no-table': (row_num, row_num)})
        start_at_row = row_num + 1

    # there may be trailing rows after the last out-of-cell content row, they will merge into a table
    if start_at_row <= worksheet_rows:
        row_segments.append({'table': (start_at_row, worksheet_rows)})

    # we have got the segments, now we render them - if table, we render as table, if no-table, we render into the doc
    segment_count = 0
    for row_segment in row_segments:
        segment_count = segment_count + 1
        if 'table' in row_segment:
            # debug('table segment {0}/{1} : spanning rows [{2}:{3}]'.format(segment_count, len(row_segments), row_segment['table'][0], row_segment['table'][1]))
            insert_content_as_table(data=data, doc=doc, start_row=start_row, start_col=start_col, row_from=row_segment['table'][0], row_to=row_segment['table'][1], container_width=container_width, container=container, cell=cell, repeat_rows=repeat_rows)
        elif 'no-table' in row_segment:
            # debug('no-table segment {0}/{1} : at row [{2}]'.format(segment_count, len(row_segments), row_segment['no-table'][0]))
            insert_content_into_doc(data=data, doc=doc, start_row=start_row, row_from=row_segment['no-table'][0], container_width=container_width)
        else:
            warn('something unsual happened - unknown row segment type')


def insert_content_into_doc(data, doc, start_row, row_from, container_width):
    # the content is by default one row content and we are only interested in the first column value
    row_data = data['sheets'][0]['data'][0]['rowData']
    first_cell_data = row_data[row_from - (start_row + 1)]['values'][0]

    # thre may be two cases
    # the value may have a 'contents' object in which case we call insert_content
    if 'contents' in first_cell_data:
        # Hack: we put a blank small-height paragraph so that it does not get merged with any previous table
        paragraph = doc.add_paragraph()
        paragraph.style = 'Calibri-2-Gray8'
        insert_content(first_cell_data['contents'], doc, container_width, container=None, cell=None)
        # polish_table(table)

    # or it may be anything else
    else:
        render_content_in_doc(doc, first_cell_data)


def insert_content_as_table(data, doc, start_row, start_col, row_from, row_to, container_width, container=None, cell=None, repeat_rows=0):
    start_time = int(round(time.time() * 1000))
    current_time = int(round(time.time() * 1000))
    last_time = current_time

    # calculate table dimension
    table_rows = row_to - row_from + 1
    table_cols = data['sheets'][0]['properties']['gridProperties']['columnCount'] - start_col

    merge_data = {}
    if 'merges' in data['sheets'][0]:
        merge_data = data['sheets'][0]['merges']

    # create the table
    if container is not None:
        table = container.add_table(table_rows, table_cols, Pt(container_width))

	# table to be added inside a cell
    elif cell is not None:
		# insert the table in the very first paragraph of the cell
        # debug('embedding new table inside an existing cell, size: {0} x {1}'.format(table_rows, table_cols))
        cell.paragraphs[0].style = 'Calibri-2-Gray8'
        table = cell.add_table(table_rows, table_cols)
    else:
        # debug('embedding new table inside the doc, size: {0} x {1}'.format(table_rows, table_cols))
        table = doc.add_table(table_rows, table_cols)

    # resize columns as per data
    column_data = data['sheets'][0]['data'][0]['columnMetadata']
    total_width = sum(x['pixelSize'] for x in column_data)
    column_widths = [ (x['pixelSize'] * container_width / total_width) for x in column_data ]

    # if the table had too many columns, use a style where there is smaller left, right margin
    if len(column_widths) > 10:
        table.style = 'PlainTable'

    last_time = current_time

    # populate cells
    # total_rows = len(data['sheets'][0]['data'][0]['rowData'])
    total_rows = table_rows
    i = 0
    current_time = int(round(time.time() * 1000))
    if not container: info('  .. rendering {0} rows'.format(total_rows))
    last_time = current_time

    # TODO: handle table related instructions from notes. table related instructions are given as notes in the very first cell (row 0, col 0). they may be
    # table-style - to apply style to whole table (NOT IMPLEMENTED YET)
    # table-spacing - no-spacing means cell paragraphs must not have any spacing throughout the table, useful for source code rendering
    # table-header-rows - number of header rows to repeat across pages (NOT IMPLEMENTED YET)

    row_data = data['sheets'][0]['data'][0]['rowData']

    # get the first cell notes
    # if there is a note, see if it is a JSON, it may contain table specific styling directives
    first_cell_data = row_data[row_from - (start_row + 1)]['values'][0]
    first_cell_note_json = {}
    if 'note' in first_cell_data:
        try:
            first_cell_note_json = json.loads(first_cell_data['note'])
        except json.JSONDecodeError:
            pass

    # handle table-spacing in notes, if the value is no-spacing then all cell paragraphs must have no spacing
    if 'table-spacing' in first_cell_note_json:
        table_spacing = first_cell_note_json['table-spacing']
        # debug('table-spacing is [{0}]'.format(table_spacing))
    else:
        table_spacing = ''

    repeating_row_count = 0
    # handle repeat-rows directive. The value is an integer telling us how many rows (from the first row) should be repeated in pages for this table
    if 'repeat-rows' in first_cell_note_json:
        repeating_row_count = int(first_cell_note_json['repeat-rows'])
        # debug('repeat-rows is [{0}]'.format(repeating_row_count))
    else:
        repeating_row_count = 0

    table_row_index = 0
    for data_row_index in range(row_from - (start_row + 1), row_to - (start_row + 0)):
        if 'values' in row_data[data_row_index]:
            row = table.row_cells(table_row_index)
            row_values = row_data[data_row_index]['values']

            for c in range(0, len(row_values)):
                # render_content_in_cell () is the main work function for rendering an individual cell (eg., gsheet cell -> docx table cell)
                render_content_in_cell(doc, row[c], row_values[c], column_widths[c], data_row_index, c, start_row, start_col, merge_data, column_widths, table_spacing)

            if table_row_index % 100 == 0:
                current_time = int(round(time.time() * 1000))
                if not container: info('  .... cell rendered for {0}/{1} rows : {2} ms'.format(table_row_index, total_rows, current_time - last_time))
                last_time = current_time

        table_row_index = table_row_index + 1

    current_time = int(round(time.time() * 1000))
    if not container: info('  .. rendering cell complete for {0} rows : {1} ms\n'.format(total_rows, current_time - start_time))
    last_time = current_time

    # merge cells according to data
    if not container: info('  .. merging cells'.format(current_time - last_time))
    for m in merge_data:
        # check if the merge is applicable for this table's rows
        if m['startRowIndex'] < (row_from - 1) or m['endRowIndex'] > (row_to):
            continue

        start_row_index = m['startRowIndex'] - (row_from - start_row) - 1
        end_row_index = m['endRowIndex'] - (row_from - start_row) - 2

        start_column_index = m['startColumnIndex'] - start_col
        end_column_index = m['endColumnIndex'] - start_col - 1
        # debug('merging cell ({0}, {1}) with cell ({2}, {3})'.format(start_row_index, start_column_index, end_row_index, end_column_index))
        starting_cell = table.cell(start_row_index, start_column_index)

        # all cells within the merge range need to have the same border as the first cell
        for r in range(start_row_index, end_row_index + 1):
            for c in range(start_column_index, end_column_index + 1):
                if (r, c) != (start_row_index, start_column_index):
                    to_cell = table.cell(r, c)
                    # TODO: copy border to all cells
                    copy_cell_border(starting_cell, to_cell)

        ending_cell = table.cell(end_row_index, end_column_index)
        starting_cell.merge(ending_cell)

    # handle repeat_rows
    for r in range(0, repeating_row_count):
        # debug('repeating row : {0}'.format(repeating_row_count))
        set_repeat_table_header(table.rows[r])

    current_time = int(round(time.time() * 1000))
    if not container: info('  .. cells merged : {0} ms\n'.format(current_time - last_time))
    last_time = current_time

    if not container: info('.. content insertion completed : {0} ms\n'.format(current_time - start_time))

    return table


def merged_cell_width(row, col, start_row, start_col, merges, column_widths):
    cell_width = 0
    for merge in merges:
        if merge['startRowIndex'] == (row + start_row) and merge['startColumnIndex'] == (col + start_col):
            for c in range(col, merge['endColumnIndex'] - start_col):
                cell_width = cell_width + column_widths[c]

    if cell_width == 0:
        return column_widths[col]
    else:
        return cell_width


def set_header(doc, section, header_first, header_odd, header_even, actual_width, linked_to_previous=False):
    first_page_header = section.first_page_header
    odd_page_header = section.header
    even_page_header = section.even_page_header

    section.first_page_header.is_linked_to_previous = linked_to_previous
    section.header.is_linked_to_previous = linked_to_previous
    section.even_page_header.is_linked_to_previous = linked_to_previous

    if len(first_page_header.tables) == 0:
        if header_first is not None: insert_content(header_first, doc, actual_width, container=first_page_header, cell=None)

    if len(odd_page_header.tables) == 0:
        if header_odd is not None: insert_content(header_odd, doc, actual_width, container=odd_page_header, cell=None)

    if len(even_page_header.tables) == 0:
        if header_even is not None: insert_content(header_even, doc, actual_width, container=even_page_header, cell=None)


def set_footer(doc, section, footer_first, footer_odd, footer_even, actual_width, linked_to_previous=False):
    first_page_footer = section.first_page_footer
    odd_page_footer = section.footer
    even_page_footer = section.even_page_footer

    section.first_page_footer.is_linked_to_previous = linked_to_previous
    section.footer.is_linked_to_previous = linked_to_previous
    section.even_page_footer.is_linked_to_previous = linked_to_previous

    if len(first_page_footer.tables) == 0:
        if footer_first is not None: insert_content(footer_first, doc, actual_width, container=first_page_footer, cell=None)

    if len(odd_page_footer.tables) == 0:
        if footer_odd is not None: insert_content(footer_odd, doc, actual_width, container=odd_page_footer, cell=None)

    if len(even_page_footer.tables) == 0:
        if footer_even is not None: insert_content(footer_even, doc, actual_width, container=even_page_footer, cell=None)


def add_section(doc, section_data, section_spec, use_existing=False):
    if section_spec['break'] == 'CONTINUOUS':
        # if it is the only section, do not add, just get the last (current) section
        if use_existing:
            section = doc.sections[-1]
        else:
            section = doc.add_section(WD_SECTION.CONTINUOUS)
    else:
        # if it is the only section, do not add, just get the last (current) section
        if use_existing:
            section = doc.sections[-1]
        else:
            section = doc.add_section(WD_SECTION.NEW_PAGE)

    if section_spec['orient'] == 'LANDSCAPE':
        section.orient = WD_ORIENT.LANDSCAPE
    else:
        section.orient = WD_ORIENT.PORTRAIT

    section.page_width = Inches(section_spec['page_width'])
    section.page_height = Inches(section_spec['page_height'])
    section.left_margin = Inches(section_spec['left_margin'])
    section.right_margin = Inches(section_spec['right_margin'])
    section.top_margin = Inches(section_spec['top_margin'])
    section.bottom_margin = Inches(section_spec['bottom_margin'])
    section.header_distance = Inches(section_spec['header_distance'])
    section.footer_distance = Inches(section_spec['footer_distance'])
    section.gutter = Inches(section_spec['gutter'])
    section.different_first_page_header_footer = section_data['different-first-page-header-footer']

    # get the actual width
    actual_width = section.page_width.inches - section.left_margin.inches - section.right_margin.inches - section.gutter.inches

    # set header if it is not set already
    set_header(doc, section, section_data['header-first'], section_data['header-odd'], section_data['header-even'], actual_width)

    # set footer if it is not set already
    set_footer(doc, section, section_data['footer-first'], section_data['footer-odd'], section_data['footer-even'], actual_width)

    return section
