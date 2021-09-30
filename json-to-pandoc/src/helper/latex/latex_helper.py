#!/usr/bin/env python3

import json
from helper.pandoc.pandoc_util import *
from helper.latex.latex_util import *

#   ----------------------------------------------------------------------------------------------------------------
#   gsheet wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' gsheet Cell object wrapper
'''
class Cell(object):

    ''' constructor
    '''
    def __init__(self, row_num, col_num, value, default_format, column_widths):
        self.row_num, self.col_num, self.column_widths, self.default_format = row_num, col_num, column_widths, default_format
        self.text_format_runs = []
        self.cell_width = self.column_widths[self.col_num]
        self.merge_spec = CellMergeSpec()

        if value:
            self.formatted_value = value.get('formattedValue')
            self.user_entered_value = CellValue(value.get('userEnteredValue'), self.formatted_value)
            self.effective_value = CellValue(value.get('effectiveValue'))
            self.user_entered_format = CellFormat(value.get('userEnteredFormat'))
            self.effective_format = CellFormat(value.get('effectiveFormat'), self.default_format)
            for text_format_run in value.get('textFormatRuns', []):
                self.text_format_runs.append(TextFormatRun(text_format_run))

            self.note = CellNote(value.get('note'))

        else:
            # value can have a special case it can be an empty ditionary when the cell is an inner cell of a column merge
            self.merge_spec.multi_col = MultiSpan.InnerCell
            self.user_entered_value = None
            self.effective_value = None
            self.formatted_value = None
            self.user_entered_format = None
            self.effective_format = None
            self.note = None


    ''' latex code for cell content
    '''
    def content_latex(self):
        # textFormatRuns first
        if len(self.text_format_runs):
            run_value_list = []
            for text_format_run in self.text_format_runs:
                run_value_list.append(text_format_run.to_latex())

            cell_value = ''.join(run_value_list)

            # TODO: for now we just show the formatted text
            cell_value = self.user_entered_value.to_latex(self.cell_width)

        # userEnteredValue next, it can be either image or text
        elif self.user_entered_value:
            # if image, userEnteredValue will have an image
            # if text, formattedValue (which we have already included into userEnteredValue) will contain the text
            cell_value = self.user_entered_value.to_latex(self.cell_width)

        # there is a 3rd possibility, the cell has no values at all, quite an empty cell
        else:
            cell_value = f"{{}}"


        # cell halign
        if self.effective_format:
            halign = self.effective_format.halign.halign
            bgcolor = self.effective_format.bgcolor.to_latex()
        else:
            halign = HALIGN.get('LEFT')
            bgcolor = self.default_format.bgcolor.to_latex()

        # finally build the cell content
        cell_content = f"{halign} {bgcolor} {cell_value}"

        return cell_content


    ''' generates the latex code
    '''
    def to_latex(self):
        latex_lines = []

        latex_lines.append(f"% {self.merge_spec.to_string()}")

        # the cell could be an inner or last cell in a multicolumn setting
        if self.merge_spec.multi_col in [MultiSpan.InnerCell, MultiSpan.LastCell]:
            # we simply do not generate anything
            return latex_lines

        # first we go for multicolumn, multirow and column width part
        if self.effective_format:
            cell_col_span = f"\\mc{{{self.merge_spec.col_span}}}{{{self.effective_format.valign.valign}{{{self.cell_width}in}}}}"


        # next we build the cell content
        cell_content = self.content_latex()

        # finally we build the whole cell
        latex_lines.append(f"{cell_col_span} {{{cell_content}}}")

        return latex_lines


''' Cell Merge spec wrapper
'''
class CellMergeSpec(object):
    def __init__(self):
        self.multi_col = MultiSpan.No
        self.multi_row = MultiSpan.No

        self.col_span = 1
        self.row_span = 1

    def to_string(self):
        return f"multicolumn: {self.multi_col}, multirow: {self.multi_row}"


''' gsheet Row object wrapper
'''
class Row(object):

    ''' constructor
    '''
    def __init__(self, row_num, row_data, default_format, section_width, column_widths):
        self.row_num, self.section_width, self.column_widths, self.default_format = row_num, section_width, column_widths, default_format

        self.cells = []
        c = 0
        for value in row_data.get('values', []):
            self.cells.append(Cell(self.row_num, c, value, self.default_format, self.column_widths))
            c = c + 1


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

    ''' generates the top border
    '''
    def top_border(self):
        return f"\\hhline{{-----}}"


    ''' generates the bottom border
    '''
    def bottom_border(self):
        return f"\\hhline{{-----}}"


    ''' generates the latex code
    '''
    def to_latex(self):
        row_lines = []

        for cell in self.cells:
            # top border
            row_lines.append(self.top_border())

            # cells
            row_line = cell.to_latex()
            row_lines = row_lines + row_line

            row_line.append(f"\\tabularnewline\n")

            # bottom border
            row_lines.append(self.bottom_border())

            if len(row_line) > 1:
                row_lines.append('&')

        # we have an extra & as the last line
        if len(row_lines) > 0:
            row_lines.pop(-1)

        return row_lines


''' gsheet rowMetadata object wrapper
'''
class RowMetadata(object):

    ''' constructor
    '''
    def __init__(self, row_metadata_dict):
        self.pixel_size = int(row_metadata_dict['pixelSize'])


''' gsheet columnMetadata object wrapper
'''
class ColumnMetadata(object):

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
        self.start_col = int(gsheet_merge_dict['startColumnIndex']) - start_column
        self.end_col = int(gsheet_merge_dict['endColumnIndex']) - start_column

        self.row_span = self.end_row - self.start_row
        self.col_span = self.end_col - self.start_col


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


    ''' generates the latex code
    '''
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


    ''' generates the latex code
    '''
    def to_latex(self):
        latex = ''
        return latex


''' gsheet cell value object wrapper
'''
class CellValue(object):

    ''' constructor
    '''
    def __init__(self, value_dict, formatted_value=None):
        if value_dict:
            if formatted_value:
                self.string_value = formatted_value
            else:
                self.string_value = value_dict.get('stringValue')

            self.image = value_dict.get('image')
        else:
            self.string_value = None
            self.image = None


    ''' generates the latex code
    '''
    def to_latex(self, cell_width):
        # if image
        if self.image:
            # even now the width may exceed actual cell width, we need to adjust for that
            dpi_x = 150 if self.image['dpi'][0] == 0 else self.image['dpi'][0]
            dpi_y = 150 if self.image['dpi'][1] == 0 else self.image['dpi'][1]
            image_width = self.image['width'] / dpi_x
            image_height = self.image['height'] / dpi_y
            if image_width > cell_width:
                adjust_ratio = (cell_width / image_width)
                # keep a padding of 0.1 inch
                image_width = cell_width - 0.2
                image_height = image_height * adjust_ratio

            latex = f"{{\includegraphics[width={image_width}in]{{{os_specific_path(self.image['path'])}}}}}"

        # if text, formattedValue will contain the text
        else:
            text = self.string_value
            if text is None:
                text = ''

            latex = f"{{{tex_escape(text)}}}"

        return latex


''' gsheet cell format object wrapper
'''
class CellFormat(object):

    ''' constructor
    '''
    def __init__(self, format_dict, default_format=None):
        if format_dict:
            self.bgcolor = RgbColor(format_dict.get('backgroundColor'))
            self.borders = Borders(format_dict.get('borders'))
            self.padding = Padding(format_dict.get('padding'))
            self.halign = HorizontalAlignment(format_dict.get('horizontalAlignment'))
            self.valign = VerticalAlignment(format_dict.get('verticalAlignment'))
            self.text_format = TextFormat(format_dict.get('textFormat'))
        elif default_format:
            self.bgcolor = default_format.bgcolor
            self.borders = default_format.borders
            self.padding = default_format.padding
            self.halign = default_format.halign
            self.valign = default_format.valign
            self.text_format = default_format.text_format
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


    ''' generates the latex code
    '''
    def to_latex(self):
        # TODO: for now returning something fixed
        latex = f"{{Missing: Custom Formatted Text Here}}"

        return latex


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

            self.header_rows = int(note_dict.get('repeat-rows', 0))
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

''' Latex section base object
'''
class LatexSectionBase(object):

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

        self.default_format = None

        if section_contents:
            self.has_content = True

            properties = section_contents.get('properties')
            self.default_format = CellFormat(properties.get('defaultFormat'))

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

                    # rowMetadata
                    for row_metadata in data.get('rowMetadata', []):
                        self.row_metadata_list.append(RowMetadata(row_metadata))

                    # columnMetadata
                    for column_metadata in data.get('columnMetadata', []):
                        self.column_metadata_list.append(ColumnMetadata(column_metadata))

                    # merges
                    for merge in sheets[0].get('merges', []):
                        self.merge_list.append(Merge(merge, self.start_row, self.start_column))

                    # column width needs adjustment as \tabcolsep is COLSEPin. This means each column has a COLSEP inch on left and right as space which needs to be removed from column width
                    all_column_widths_in_pixel = sum(x.pixel_size for x in self.column_metadata_list)
                    self.column_widths = [ (x.pixel_size * self.section_width / all_column_widths_in_pixel) - (COLSEP * 2) for x in self.column_metadata_list ]

                    # rowData
                    r = 0
                    for row_data in data.get('rowData', []):
                        self.cell_matrix.append(Row(r, row_data, self.default_format, self.section_width, self.column_widths))
                        r = r + 1

        else:
            self.has_content = False

        # we need a list to hold the tables and block for the cells
        self.content_list = []

        # generate the header block
        header_block = LatexSectionHeader(self.section_spec, self.section_break, self.level, self.no_heading, self.section, self.heading, self.title)
        self.content_list.append(header_block)


    ''' processes the cells to generate the proper order of tables and blocks
    '''
    def process(self):
        pass


    ''' generates the latex code
    '''
    def to_latex(self):
        return ''


''' Latex section object
'''
class LatexSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, section_spec):
        super().__init__(section_data, section_spec)


    ''' processes the cells to generate the proper order of tables and blocks
    '''
    def process(self):

        # update the cells' CellMergeSpec
        for merge in self.merge_list:
            first_row = merge.start_row
            first_col = merge.start_col
            last_row = merge.end_row
            last_col = merge.end_col
            cell = self.cell_matrix[first_row].get_cell(first_col)
            if cell:
                cell.merge_spec.multi_col = MultiSpan.FirstCell
                cell.merge_spec.multi_row = MultiSpan.FirstCell

                cell.merge_spec.col_span = merge.col_span
                cell.merge_spec.row_span = merge.row_span

                for c in range(first_col + 1, last_col):
                    cell.cell_width = cell.cell_width + self.column_widths[c] + COLSEP * 2


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
                    table = LatexTable(self.cell_matrix, next_table_starts_in_row, r - 1, self.column_widths)
                    self.content_list.append(table)

                block = LatexParagraph(data_row, r)
                self.content_list.append(block)

                next_table_starts_in_row = r + 1

            else:
                next_table_ends_in_row = r

        # there may be a pending/running table
        if next_table_ends_in_row >= next_table_starts_in_row:
            table = LatexTable(self.cell_matrix, next_table_starts_in_row, next_table_ends_in_row, self.column_widths)
            self.content_list.append(table)


    ''' generates the latex code
    '''
    def to_latex(self):
        # process first
        self.process()

        latex_lines = []

        # iterate to through tables and blocks contents
        for block in self.content_list:
            latex_lines = latex_lines + block.to_latex()

        return latex_lines


''' Latex section object
'''
class LatexToCSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, section_spec):
        super().__init__(section_data, section_spec)


    ''' processes the cells to generate the proper order of tables and blocks
    '''
    def process(self):
        pass

    ''' generates the latex code
    '''
    def to_latex(self):
        # process first
        self.process()

        latex_lines = []

        # iterate to through tables and blocks contents
        for block in self.content_list:
            latex_lines = latex_lines + block.to_latex()

        return latex_lines


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
    def __init__(self, cell_matrix, start_row, end_row, column_widths):
        self.start_row, self.end_row, self.column_widths = start_row, end_row, column_widths
        self.table_cell_matrix = cell_matrix[start_row:end_row+1]
        self.row_count = len(self.table_cell_matrix)

        # header row if any
        self.header_row_count = self.table_cell_matrix[0].get_cell(0).note.header_rows

    ''' generates the latex code
    '''
    def to_latex(self):
        table_col_spec = ' '.join([f"p{{{i}in}}" for i in self.column_widths])
        table_lines = []

        table_lines.append(begin_latex())
        table_lines.append(f"% LatexTable: ({self.start_row}-{self.end_row}) : {self.row_count} rows")
        table_lines.append(f"\\setlength\\parindent{{0pt}}")
        table_lines.append(f"\\begin{{longtable}}[l]{{|{table_col_spec}|}}\n")

        # TODO: generate the table
        r = 1
        for row in self.table_cell_matrix:
            table_lines = table_lines + row.to_latex()

            # header row
            if self.header_row_count == r:
                table_lines.append(f"\\endhead\n")

            r = r + 1

        table_lines.append(f"\n\\end{{longtable}}")
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
        if len(self.data_row.cells) > 0:
            row_text = self.data_row.get_cell(0).content_latex()
            block_lines.append(row_text)

        block_lines.append(end_latex())
        return block_lines
