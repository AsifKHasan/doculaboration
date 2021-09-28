#!/usr/bin/env python3

import json
from helper.latex.latex_util import *

#   ----------------------------------------------------------------------------------------------------------------
#   gsheet wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' gsheet Cell object wrapper
'''
class Cell(object):

    ''' constructor
    '''
    def __init__(self, value):
        self.text_format_runs = []
        if value:
            self.user_entered_value = CellValue(value.get('userEnteredValue'))
            self.effective_value = CellValue(value.get('effectiveValue'))
            self.formatted_value = value.get('formattedValue')
            self.user_entered_format = CellFormat(value.get('userEnteredFormat'))
            self.effective_format = CellFormat(value.get('effectiveFormat'))
            for text_format_run in value.get('textFormatRuns', []):
                self.text_format_runs.append(TextFormatRun(text_format_run))

            self.note = CellNote(value.get('note'))

        else:
            self.user_entered_value = None
            self.effective_value = None
            self.formatted_value = None
            self.user_entered_format = None
            self.effective_format = None
            self.note = None


''' gsheet Row object wrapper
'''
class Row(object):

    ''' constructor
    '''
    def __init__(self, row_data):
        self.cells = []
        for value in row_data.get('values', []):
            self.cells.append(Cell(value))

    def is_empty(self):
        return (len(self.cells) == 0)

    def get_cell(self, c):
        if c >= 0 and c < len(self.cells):
            return self.cells[c]
        else:
            return None

    def is_out_of_table(self):
        if len(self.cells) > 0:
            # the first cell is teh relevant cell only
            return self.cells[0].note.out_of_table
        else:
            return False

''' gsheet rowMetaData object wrapper
'''
class RowMetaData(object):

    ''' constructor
    '''
    def __init__(self, row_metadata_dict):
        self.pixel_size = int(row_metadata_dict['pixelSize'])


''' gsheet columnMetaData object wrapper
'''
class ColumnMetaData(object):

    ''' constructor
    '''
    def __init__(self, column_metadata_dict):
        self.pixel_size = int(column_metadata_dict['pixelSize'])


''' gsheet merge object wrapper
'''
class Merge(object):

    ''' constructor
    '''
    def __init__(self, gsheet_merge_dict, start_row, start_column):
        self.start_row = int(gsheet_merge_dict['startRowIndex']) - start_row
        self.end_row = int(gsheet_merge_dict['endRowIndex']) - start_row
        self.start_column = int(gsheet_merge_dict['startColumnIndex']) - start_column
        self.end_column = int(gsheet_merge_dict['endColumnIndex']) - start_column

        self.row_span = self.end_row - self.start_row
        self.column_span = self.end_column - self.start_column


''' gsheet color object wrapper
'''
class RgbColor(object):

    ''' constructor
    '''
    def __init__(self, rgb_dict=None):
        if rgb_dict:
            self.red = float(rgb_dict.get('red', 0))
            self.green = float(rgb_dict.get('green', 0))
            self.blue = float(rgb_dict.get('blue', 0))
        else:
            self.red = 0
            self.green = 0
            self.blue = 0

    def to_latex(self):
        return f"\cellcolor[rgb]{{{self.red},{self.green},{self.blue}}}"


''' gsheet cell padding object wrapper
'''
class Padding(object):

    ''' constructor
    '''
    def __init__(self, padding_dict=None):
        if padding_dict:
            self.top = int(padding_dict.get('top', 0))
            self.right = int(padding_dict.get('right', 0))
            self.bottom = int(padding_dict.get('bottom', 0))
            self.left = int(padding_dict.get('left', 0))
        else:
            self.top = 0
            self.right = 0
            self.bottom = 0
            self.left = 0


''' gsheet text format object wrapper
'''
class TextFormat(object):

    ''' constructor
    '''
    def __init__(self, text_format_dict=None):
        if text_format_dict:
            self.fgcolor = RgbColor(text_format_dict.get('foregroundColor'))
            self.font_family = text_format_dict.get('fontFamily')
            self.font_size = int(text_format_dict.get('fontSize', 0))
            self.is_bold = text_format_dict.get('bold')
            self.is_italic = text_format_dict.get('italic')
            self.is_strikethrough = text_format_dict.get('strikethrough')
            self.is_underline = text_format_dict.get('underline')
        else:
            self.fgcolor = RgbColor()
            self.font_family = None
            self.font_size = 0
            self.is_bold = False
            self.is_italic = False
            self.is_strikethrough = False
            self.is_underline = False


''' gsheet cell borders object wrapper
'''
class Borders(object):

    ''' constructor
    '''
    def __init__(self, borders_dict=None):
        if borders_dict:
            self.top = Border(borders_dict.get('top'))
            self.right = Border(borders_dict.get('right'))
            self.bottom = Border(borders_dict.get('bottom'))
            self.left = Border(borders_dict.get('left'))
        else:
            self.top = None
            self.right = None
            self.bottom = None
            self.left = None


''' gsheet cell border object wrapper
'''
class Border(object):

    ''' constructor
    '''
    def __init__(self, border_dict=None):
        if border_dict:
            self.style = border_dict.get('style')
            self.width = int(border_dict.get('width'))
            self.color = RgbColor(border_dict.get('color'))
        else:
            self.style = None
            self.width = None
            self.color = None


''' gsheet cell value object wrapper
'''
class CellValue(object):

    ''' constructor
    '''
    def __init__(self, value_dict):
        if value_dict:
            self.string_value = value_dict.get('stringValue')
            self.image = value_dict.get('image')
        else:
            self.string_value = None
            self.image = None


''' gsheet cell format object wrapper
'''
class CellFormat(object):

    ''' constructor
    '''
    def __init__(self, format_dict):
        if format_dict:
            self.bgcolor = RgbColor(format_dict.get('backgroundColor'))
            self.borders = Borders(format_dict.get('borders'))
            self.padding = Padding(format_dict.get('padding'))
            self.halign = HorizontalAlignment(format_dict.get('horizontalAlignment'))
            self.valign = VerticalAlignment(format_dict.get('verticalAlignment'))
            self.text_format = TextFormat(format_dict.get('textFormat'))
        else:
            self.bgcolor = None
            self.borders = None
            self.padding = None
            self.halign = None
            self.valign = None
            self.text_format = None


''' gsheet text format run object wrapper
'''
class TextFormatRun(object):

    ''' constructor
    '''
    def __init__(self, run_dict=None):
        if run_dict:
            self.start_index = int(run_dict.get('startIndex', 0))
            self.format = TextFormat(run_dict.get('format'))
        else:
            self.start_index = None
            self.format = None


''' gsheet cell notes object wrapper
'''
class CellNote(object):

    ''' constructor
    '''
    def __init__(self, note_json=None):
        self.style = None
        self.out_of_table = False
        self.table_spacing = True
        self.header_rows = 0
        self.new_page = False
        self.keep_with_next = False
        self.page_number_style = None

        if note_json:
            try:
                note_dict = json.loads(note_json)
            except json.JSONDecodeError:
                note_dict = {}

            self.style = note_dict.get('style')

            content = note_dict.get('content')
            if content is not None and content == 'out-of-cell':
                self.out_of_table = True

            spacing = note_dict.get('table-spacing')
            if spacing is not None and spacing == 'no-spacing':
                self.table_spacing = False

            self.header_rows = int(note_dict.get('header-rows', 0))
            self.new_page = note_dict.get('new-page') is not None
            self.keep_with_next = note_dict.get('keep-with-next') is not None
            self.page_number_style = note_dict.get('page-number')


''' gsheet vertical alignment object wrapper
'''
class VerticalAlignment(object):

    ''' constructor
    '''
    def __init__(self, valign=None):
        if valign:
            self.valign = VALIGN.get(valign, 'p')
        else:
            self.valign = VALIGN.get('TOP')


''' gsheet horizontal alignment object wrapper
'''
class HorizontalAlignment(object):

    ''' constructor
    '''
    def __init__(self, halign=None):
        if halign:
            self.halign = HALIGN.get(halign, 'LEFT')
        else:
            self.halign = HALIGN.get('LEFT')


#   ----------------------------------------------------------------------------------------------------------------
#   latex wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' Latex section object wrapper
'''
class LatexSection(object):

    ''' constructor
    '''
    def __init__(self, section_data, section_spec):
        self.section_spec = section_spec
        self.section_break = section_data['section-break']
        self.section_width = float(self.section_spec['page_width']) - float(self.section_spec['left_margin']) - float(self.section_spec['right_margin']) - float(self.section_spec['gutter'])

        self.no_heading = section_data['no-heading']
        self.section = section_data['section']
        self.heading = section_data['heading']
        self.level = section_data['level']
        self.page_numbering = section_data['page-number']

        section_contents = section_data.get('contents')

        self.title = None
        self.row_count = 0
        self.column_count = 0

        self.start_row = 0
        self.start_column = 0

        self.cell_matrix = []
        self.row_metadata_list = []
        self.column_metadata_list = []
        self.merge_list = []

        self.default_bgolor = RgbColor()
        self.default_padding = Padding()
        self.default_valign = None
        self.default_format = None

        if section_contents:
            self.has_content = True
            properties = section_contents.get('properties')
            if properties and 'defaultFormat' in properties:
                self.default_bgolor = RgbColor(properties['defaultFormat'].get('backgroundColor'))
                self.default_padding = Padding(properties['defaultFormat'].get('padding'))
                self.default_valign = VerticalAlignment(properties['defaultFormat'].get('verticalAlignment'))
                self.default_format = TextFormat(properties['defaultFormat'].get('textFormat'))

            sheets = section_contents.get('sheets')
            if isinstance(sheets, list) and len(sheets) > 0:
                sheet_properties = sheets[0].get('properties')
                if sheet_properties:
                    self.title = sheet_properties.get('title')
                    if 'gridProperties' in sheet_properties:
                        self.row_count = max(int(sheet_properties['gridProperties'].get('rowCount', 0)) - 2, 0)
                        self.column_count = max(int(sheet_properties['gridProperties'].get('columnCount', 0)) - 1, 0)

                data_list = sheets[0].get('data')
                if isinstance(data_list, list) and len(data_list) > 0:
                    data = data_list[0]
                    self.start_row = int(data.get('startRow', 0))
                    self.start_column = int(data.get('startColumn', 0))

                    # rowData
                    for row_data in data.get('rowData', []):
                        self.cell_matrix.append(Row(row_data))

                    # rowMetaData
                    for row_metadata in data.get('rowMetaData', []):
                        self.row_metadata_list.append(RowMetaData(row_metadata))

                    # columnMetaData
                    for column_metadata in data.get('columnMetaData', []):
                        self.column_metadata_list.append(ColumnMetaData(column_metadata))

                    # merges
                    for merge in sheets[0].get('merges', []):
                        self.merge_list.append(Merge(merge, self.start_row, self.start_column))

        else:
            self.has_content = False

        # we need a list to hold the tables and block for the cells
        self.content_list = []
        self.process()


    ''' processes the cells to generate the proper order of tables and blocks
    '''
    def process(self):
        # header
        header_block = LatexSectionHeader(self.section_spec, self.section_break, self.level, self.no_heading, self.section, self.heading, self.title)
        self.content_list.append(header_block)

        # we have a concept of in-cell content and out-of-cell content
        # in-cell contents are treated as part of a table, while out-of-cell contents are treated as independent paragraphs, images etc. (blocks)
        next_table_starts_in_row = 0
        next_table_ends_in_row = 0
        for r in range(0, self.row_count):
            # the first cell of the row tells us whether it is in-cell or out-of-cell
            data_row = self.cell_matrix[r]
            if data_row.is_out_of_table() == True:
                # there may be a pending/running table
                if r > next_table_starts_in_row:
                    table = LatexTable(self.cell_matrix, next_table_starts_in_row, r - 1)
                    self.content_list.append(table)

                block = LatexParagraph(data_row, r)
                self.content_list.append(block)

                next_table_starts_in_row = r + 1

            else:
                next_table_ends_in_row = r

        # there may be a pending/running table
        if next_table_ends_in_row >= next_table_starts_in_row:
            table = LatexTable(self.cell_matrix, next_table_starts_in_row, next_table_ends_in_row)
            self.content_list.append(table)


    ''' generates the latex code
    '''
    def to_latex(self):
        latex_lines = []

        # iterate to through tables and blocks contents
        for block in self.content_list:
            latex_lines = latex_lines + block.to_latex()

        return '\n'.join(latex_lines)


''' Latex Block object wrapper base class (plain latex, table, header etc.)
'''
class LatexBlock(object):

    ''' generates latex code
    '''
    def to_latex(self):
        pass


''' Latex Header Block wrapper
'''
class LatexSectionHeader(LatexBlock):

    ''' constructor
    '''
    def __init__(self, section_spec, section_break, level, no_heading, section, heading, title):
        self.section_spec, self.section_break, self.level, self.no_heading, self.section, self.heading, self.title = section_spec, section_break, level, no_heading, section, heading, title

    ''' generates latex code
    '''
    def to_latex(self):
        header_lines = []
        header_lines.append('')
        header_lines.append(begin_latex())
        header_lines.append(f"% LatexSection: {self.title}")

        if self.section_break.startswith('newpage_'):
            header_lines.append(f"\\newpage")

        header_lines.append(f"\pdfpagewidth {self.section_spec['page_width']}in")
        header_lines.append(f"\pdfpageheight {self.section_spec['page_height']}in")
        header_lines.append(f"\\newgeometry{{top={self.section_spec['top_margin']}in, bottom={self.section_spec['bottom_margin']}in, left={self.section_spec['left_margin']}in, right={self.section_spec['right_margin']}in}}")

        header_lines.append(end_latex())

        # heading
        if not self.no_heading:
            if self.section != '':
                heading_text = f"{'#' * self.level} {self.section} - {self.heading}".strip()
            else:
                heading_text = f"{'#' * self.level} {self.heading}".strip()

            # headings are styles based on level
            if self.level != 0:
                header_lines.append(heading_text)
                header_lines.append('\n')
            else:
                header_lines.append(begin_latex())
                header_lines.append(f"\\titlestyle{{{heading_text}}}")
                header_lines.append(end_latex())

        return header_lines


''' Latex Table object wrapper
'''
class LatexTable(LatexBlock):

    ''' constructor
    '''
    def __init__(self, cell_matrix, start_row, end_row):
        self.start_row, self.end_row = start_row, end_row
        self.table_cell_matrix = cell_matrix[start_row:end_row+1]
        self.row_count = len(self.table_cell_matrix)

    ''' generates the latex code
    '''
    def to_latex(self):
        table_lines = []
        table_lines.append(begin_latex())
        table_lines.append(f"% LatexTable: ({self.start_row}-{self.end_row}) : {self.row_count} rows")

        # TODO: generate the table

        table_lines.append(end_latex())
        return table_lines


''' Latex Block object wrapper
'''
class LatexParagraph(LatexBlock):

    ''' constructor
    '''
    def __init__(self, data_row, row_number):
        self.data_row = data_row
        self.row_number = row_number

    ''' generates the latex code
    '''
    def to_latex(self):
        block_lines = []
        block_lines.append(begin_latex())
        block_lines.append(f"% LatexParagraph: row {self.row_number}")

        # TODO: generate the block

        block_lines.append(end_latex())
        return block_lines
