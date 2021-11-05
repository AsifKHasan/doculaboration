#!/usr/bin/env python3

import json
from helper.pandoc.pandoc_util import *
from helper.latex.latex_util import *

#   ----------------------------------------------------------------------------------------------------------------
#   gsheet cell wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' gsheet Cell object wrapper
'''
class Cell(object):

    ''' constructor
    '''
    def __init__(self, row_num, col_num, value, default_format, column_widths):
        self.row_num, self.col_num, self.column_widths, self.default_format = row_num, col_num, column_widths, default_format
        self.cell_name = f"cell: [{self.row_num},{self.col_num}]"
        self.value = value
        self.text_format_runs = []
        self.cell_width = self.column_widths[self.col_num]
        self.merge_spec = CellMergeSpec()

        if self.value:
            self.formatted_value = self.value.get('formattedValue')
            self.user_entered_value = CellValue(self.value.get('userEnteredValue'), self.formatted_value)
            self.effective_value = CellValue(self.value.get('effectiveValue'))
            self.user_entered_format = CellFormat(self.value.get('userEnteredFormat'))
            self.effective_format = CellFormat(self.value.get('effectiveFormat'), self.default_format)
            for text_format_run in self.value.get('textFormatRuns', []):
                self.text_format_runs.append(TextFormatRun(text_format_run, self.effective_format.text_format.source))

            self.note = CellNote(value.get('note'))
            self.is_empty = False

            # handle special notes
            if self.note.page_number:
                self.user_entered_value.string_value = "\\thepage\\ of \\pageref{LastPage}"
                self.user_entered_value.verbatim = True

        else:
            # value can have a special case it can be an empty ditionary when the cell is an inner cell of a column merge
            self.merge_spec.multi_col = MultiSpan.No
            self.user_entered_value = None
            self.effective_value = None
            self.formatted_value = None
            self.user_entered_format = None
            self.effective_format = None
            self.note = CellNote()
            self.is_empty = True


    ''' string representation
    '''
    def __repr__(self):
        s = f"[{self.row_num+1},{self.col_num+1}], value: {not self.is_empty}, wd: {self.cell_width}in, mr: {self.merge_spec.multi_row}, mc: {self.merge_spec.multi_col}"
        if self.effective_format:
            b = f"{self.user_entered_format.borders}"
        else:
            b = f"No Border"

        return f"{s}....{b}"


    ''' Copy format from the cell passed
    '''
    def copy_format_from(self, from_cell):
        self.user_entered_format = from_cell.user_entered_format
        self.effective_format = from_cell.effective_format

        self.merge_spec.multi_col = from_cell.merge_spec.multi_col
        self.merge_spec.col_span = from_cell.merge_spec.col_span
        self.merge_spec.row_span = from_cell.merge_spec.row_span
        self.cell_width = from_cell.cell_width


    ''' adjust cell borders based on cell's bgcolor
    '''
    def adjust_borders(self):
        # if the cell has a bgcolor and any border is missing, that border should be colored as bgcolor
        if self.user_entered_format and self.user_entered_format.bgcolor:
            self.effective_format.override_borders(self.user_entered_format.bgcolor)

        # only if multi_col is No or FirstCell
        if self.merge_spec.multi_col in [MultiSpan.FirstCell, MultiSpan.No]:
            # if the cell is multi_row FirstCell or InnerCell, bottom border color is bgcolor
            if self.merge_spec.multi_row in [MultiSpan.FirstCell, MultiSpan.InnerCell]:
                self.effective_format.recolor_bottom_border(self.user_entered_format.bgcolor)

            # if the cell is multi_row InnerCell or LastCell, top border color is bgcolor
            if self.merge_spec.multi_row in [MultiSpan.InnerCell, MultiSpan.LastCell]:
                self.effective_format.recolor_top_border(self.user_entered_format.bgcolor)


    ''' mark the cell multi_col
    '''
    def mark_multicol(self, span):
        self.merge_spec.multi_col = span


    ''' mark the cell multi_col
    '''
    def mark_multirow(self, span):
        self.merge_spec.multi_row = span


    ''' whether top border is allowed for the cell
    '''
    def top_border_allowed(self):
        if self.merge_spec.multi_row in [MultiSpan.No, MultiSpan.FirstCell]:
            if self.merge_spec.multi_col in [MultiSpan.No, MultiSpan.FirstCell]:
                return True

        return False


    ''' whether bottom border is allowed for the cell
    '''
    def bottom_border_allowed(self):
        if self.merge_spec.multi_row in [MultiSpan.No, MultiSpan.LastCell]:
            if self.merge_spec.multi_col in [MultiSpan.No, MultiSpan.FirstCell]:
                return True

        return False


    ''' whether left border is allowed for the cell
    '''
    def left_border_allowed(self):
        if self.merge_spec.multi_col in [MultiSpan.No, MultiSpan.FirstCell]:
            return True

        return False


    ''' whether right border is allowed for the cell
    '''
    def right_border_allowed(self):
        if self.merge_spec.multi_col in [MultiSpan.No, MultiSpan.LastCell]:
            return True

        return False


    ''' latex code for top border
    '''
    def top_border_latex(self, color_dict):
        t = None
        if self.effective_format and self.effective_format.borders:
            if self.top_border_allowed():
                t = self.effective_format.borders.to_latex_t(color_dict)
                if t is not None:
                    # t = f"{{{self.col_num+1}-{self.col_num+self.merge_spec.col_span}}}{{{t}}}"
                    t = f"[{t}]{{{self.col_num+1}-{self.col_num+self.merge_spec.col_span}}}"

        return t


    ''' latex code for bottom border
    '''
    def bottom_border_latex(self, color_dict):
        b = None
        if self.effective_format and self.effective_format.borders:
            if self.bottom_border_allowed():
                b = self.effective_format.borders.to_latex_b(color_dict)
                if b is not None:
                    # b = f"{{{self.col_num+1}-{self.col_num+self.merge_spec.col_span}}}{{{b}}}"
                    b = f"[{b}]{{{self.col_num+1}-{self.col_num+self.merge_spec.col_span}}}"

        return b


    ''' latex code for left and right borders
        r is row numbner (1 based)
    '''
    def cell_vertical_borders_latex(self, r, color_dict):
        lr_borders = []
        if self.effective_format and self.effective_format.borders:
            if self.left_border_allowed():
                lb = self.effective_format.borders.to_latex_l(color_dict)
                if lb is not None:
                    lb = f"vline{{{self.col_num+1}}} = {{{r}}}{{{lb}}},"
                    lr_borders.append(lb)

            if self.right_border_allowed():
                rb = self.effective_format.borders.to_latex_r(color_dict)
                if rb is not None:
                    rb = f"vline{{{self.col_num+2}}} = {{{r}}}{{{rb}}},"
                    lr_borders.append(rb)

        return lr_borders


    ''' latex code for cell content
    '''
    def content_latex(self, include_formatting, color_dict, strip_comments=False):
        content_lines = []

        if not strip_comments:
            content_lines.append(f"% {self.merge_spec.to_string()}")

        # the content is not valid for multirow LastCell and InnerCell
        if self.merge_spec.multi_row in [MultiSpan.InnerCell, MultiSpan.LastCell]:
            cell_value = None

        elif self.merge_spec.multi_col in [MultiSpan.InnerCell, MultiSpan.LastCell]:
            cell_value = None

        else:
            # textFormatRuns first
            if len(self.text_format_runs):
                run_value_list = []
                processed_idx = len(self.formatted_value)
                for text_format_run in reversed(self.text_format_runs):
                    text = self.formatted_value[:processed_idx]
                    run_value_list.insert(0, text_format_run.to_latex(text, color_dict))
                    processed_idx = text_format_run.start_index

                cell_value = ''.join(run_value_list)

            # userEnteredValue next, it can be either image or text
            elif self.user_entered_value:
                # if image, userEnteredValue will have an image
                # if text, formattedValue (which we have already included into userEnteredValue) will contain the text
                cell_value = self.user_entered_value.to_latex(self.cell_width, self.effective_format.text_format, color_dict)

            # there is a 3rd possibility, the cell has no values at all, quite an empty cell
            else:
                cell_value = None

        if cell_value:
            # paragraphs need formatting to be included, table cells do not need them
            if include_formatting:
                # alignments and bgcolor
                if self.effective_format:
                    halign = PARA_HALIGN.get(self.effective_format.halign.halign)
                else:
                    halign = PARA_HALIGN.get('LEFT')

                cell_value = f"{halign}{{{cell_value}}}"

            content_lines.append(cell_value)

        return content_lines


    ''' latex code for cell format
    '''
    def format_latex(self, r, color_dict):
        latex_lines = []

        # alignments and bgcolor
        if self.effective_format:
            halign = self.effective_format.halign.halign
            valign = self.effective_format.valign.valign
            bgcolor = self.effective_format.bgcolor
        else:
            halign = TBLR_HALIGN.get('LEFT')
            valign = TBLR_VALIGN.get('MIDDLE')
            bgcolor = self.default_format.bgcolor

        # finally build the cell content
        color_dict[bgcolor.key()] = bgcolor.value()
        if not self.is_empty:
            cell_format_latex = f"cell{{{r}}}{{{self.col_num+1}}} = {{r={self.merge_spec.row_span},c={self.merge_spec.col_span}}}{{valign={valign},halign={halign},bg={bgcolor.key()},wd={self.cell_width}in}},"
            latex_lines.append(cell_format_latex)

        return latex_lines


''' gsheet Row object wrapper
'''
class Row(object):

    ''' constructor
    '''
    def __init__(self, row_num, row_data, default_format, section_width, column_widths):
        self.row_num, self.section_width, self.column_widths, self.default_format = row_num, section_width, column_widths, default_format
        self.row_name = f"row: [{self.row_num+1}]"

        self.cells = []
        c = 0
        for value in row_data.get('values', []):
            self.cells.append(Cell(self.row_num, c, value, self.default_format, self.column_widths))
            c = c + 1


    ''' is the row empty (no cells at all)
    '''
    def is_empty(self):
        return (len(self.cells) == 0)


    ''' gets a specific cell by ordinal
    '''
    def get_cell(self, c):
        if c >= 0 and c < len(self.cells):
            return self.cells[c]
        else:
            return None


    ''' inserts a specific cell at a specific ordinal
    '''
    def insert_cell(self, pos, cell):
        # if pos is greater than the last index
        if pos > len(self.cells):
            # insert None objects in between
            fill_from = len(self.cells)
            for i in range(fill_from, pos):
                self.cells.append(None)

        if pos < len(self.cells):
            self.cells[pos] = cell
        else:
            self.cells.append(cell)


    ''' it is true only when the first cell has a out_of_table true value
    '''
    def is_out_of_table(self):
        if len(self.cells) > 0:
            # the first cell is the relevant cell only
            if self.cells[0]:
                return self.cells[0].note.out_of_table
            else:
                return False
        else:
            return False


    ''' it is true only when the first cell has a repeat-rows note with value > 0
    '''
    def is_table_start(self):
        if len(self.cells) > 0:
            # the first cell is the relevant cell only
            if self.cells[0]:
                return (self.cells[0].note.header_rows > 0)
            else:
                return False
        else:
            return False


    ''' generates the top borders
    '''
    def top_borders_latex(self, color_dict):
        top_borders = []
        c = 0
        for cell in self.cells:
            if cell is None:
                warn(f"{self.row_name} has a Null cell at {c}")

            else:
                t = cell.top_border_latex(color_dict)
                if t is not None:
                    # top_borders.append(f"\\SetHline{t}")
                    top_borders.append(f"\\cline{t}")

            c = c + 1

        return top_borders


    ''' generates the bottom borders
    '''
    def bottom_borders_latex(self, color_dict):
        bottom_borders = []
        c = 0
        for cell in self.cells:
            if cell is None:
                warn(f"{self.row_name} has a Null cell at {c}")

            else:
                b = cell.bottom_border_latex(color_dict)
                if b is not None:
                    # bottom_borders.append(f"\\SetHline{b}")
                    bottom_borders.append(f"\\cline{b}")

            c = c + 1

        return bottom_borders


    ''' generates the vertical borders
    '''
    def vertical_borders_latex(self, r, color_dict):
        v_lines = []
        c = 0
        for cell in self.cells:
            if cell is not None:
                v_lines = v_lines + cell.cell_vertical_borders_latex(r, color_dict)

            c = c + 1

        return v_lines


    ''' generates the latex code for cell formats
    '''
    def cell_format_latex(self, r, color_dict):
        cell_format_lines = []
        for cell in self.cells:
            if cell is not None:
                if not cell.is_empty:
                    cell_format_lines = cell_format_lines + cell.format_latex(r, color_dict)

        return cell_format_lines


    ''' generates the latex code
    '''
    def cell_content_latex(self, include_formatting, color_dict, strip_comments=False):
        # debug(f"processing {self.row_name}")

        row_lines = []
        # if not strip_comments:
        row_lines.append(f"% {self.row_name}")

        # borders
        top_border_lines = self.top_borders_latex(color_dict)
        top_border_lines = list(map(lambda x: f"\t{x}", top_border_lines))

        bottom_border_lines = self.bottom_borders_latex(color_dict)
        bottom_border_lines = list(map(lambda x: f"\t{x}", bottom_border_lines))

        # left_border_lines = self.left_borders_latex(color_dict)
        # left_border_lines = list(map(lambda x: f"\t{x}", left_border_lines))

        # right_border_lines = self.right_borders_latex(color_dict)
        # right_border_lines = list(map(lambda x: f"\t{x}", right_border_lines))

        # gets the cell latex lines
        all_cell_lines = []
        first_cell = True
        c = 0
        for cell in self.cells:
            if cell is None:
                warn(f"{self.row_name} has a Null cell at {c}")
                cell_lines = []
            else:
                cell_lines = cell.content_latex(include_formatting=include_formatting, color_dict=color_dict, strip_comments=strip_comments)

            if c > 0:
                all_cell_lines.append('&')

            all_cell_lines = all_cell_lines + cell_lines
            c = c + 1

        all_cell_lines.append(f"\\\\")
        all_cell_lines = list(map(lambda x: f"\t{x}", all_cell_lines))


        # top border
        row_lines = row_lines + top_border_lines
        # row_lines = row_lines + left_border_lines
        # row_lines = row_lines + right_border_lines

        # all cells
        row_lines = row_lines + all_cell_lines

        # bottom border
        row_lines = row_lines + bottom_border_lines

        return row_lines


''' gsheet text format object wrapper
'''
class TextFormat(object):

    ''' constructor
    '''
    def __init__(self, text_format_dict=None):
        self.source = text_format_dict
        if self.source:
            self.fgcolor = RgbColor(text_format_dict.get('foregroundColor'))
            self.font_family = FONT_MAP.get(text_format_dict.get('fontFamily'), '')
            self.font_size = int(text_format_dict.get('fontSize', 0))
            self.is_bold = text_format_dict.get('bold')
            self.is_italic = text_format_dict.get('italic')
            self.is_strikethrough = text_format_dict.get('strikethrough')
            self.is_underline = text_format_dict.get('underline')
        else:
            self.fgcolor = RgbColor()
            self.font_family = ''
            self.font_size = 0
            self.is_bold = False
            self.is_italic = False
            self.is_strikethrough = False
            self.is_underline = False


    def to_latex(self, text, color_dict):
        content = f"{{{text}}}"
        color_dict[self.fgcolor.key()] = self.fgcolor.value()

        if self.is_underline: content = f"\\underline{{{content}}}"
        if self.is_strikethrough: content = f"\\sout{{{content}}}"
        if self.is_italic: content = f"\\textit{{{content}}}"
        if self.is_bold: content = f"\\textbf{{{content}}}"

        # color, font, font-size
        if self.font_family != '':
            fontspec = f"\\fontsize{{{self.font_size}pt}}{{{self.font_size}pt}}\\fontspec{{{self.font_family}}}\\color{{{self.fgcolor.key()}}}"
        else:
            fontspec = f"\\fontsize{{{self.font_size}pt}}{{{self.font_size}pt}}\\color{{{self.fgcolor.key()}}}"

        latex = f"{fontspec}{content}"
        return latex


''' gsheet cell value object wrapper
'''
class CellValue(object):

    ''' constructor
    '''
    def __init__(self, value_dict, formatted_value=None):
        self.verbatim = False
        if value_dict:
            if formatted_value:
                self.string_value = formatted_value
            else:
                self.string_value = value_dict.get('stringValue')

            self.image = value_dict.get('image')
        else:
            self.string_value = ''
            self.image = None


    ''' generates the latex code
    '''
    def to_latex(self, cell_width, format, color_dict):
        # if image
        if self.image:
            # even now the width may exceed actual cell width, we need to adjust for that
            dpi_x = 150 if self.image['dpi'][0] == 0 else self.image['dpi'][0]
            dpi_y = 150 if self.image['dpi'][1] == 0 else self.image['dpi'][1]
            image_width = self.image['width'] / dpi_x
            image_height = self.image['height'] / dpi_y
            debug(f"image : [{image_width}in X {image_height}in, cell-width [{cell_width}in]")
            if image_width > cell_width:
                adjust_ratio = (cell_width / image_width)
                # keep a padding of 0.1 inch
                image_width = cell_width - 0.2
                image_height = image_height * adjust_ratio

            latex = f"\includegraphics[width={image_width}in]{{{os_specific_path(self.image['path'])}}}"

        # if text, formattedValue will contain the text
        else:
            # print(self.string_value)
            if self.verbatim:
                latex = format.to_latex(self.string_value, color_dict)
            else:
                latex = format.to_latex(tex_escape(self.string_value), color_dict)

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


    ''' override borders with the specified color
    '''
    def override_borders(self, color):
        if self.borders is None:
            self.borders = Borders(None, color)
        else:
            self.borders.override_top_border(color)
            self.borders.override_bottom_border(color)
            self.borders.override_left_border(color)
            self.borders.override_right_border(color)


    ''' recolor top border with the specified color
    '''
    def recolor_top_border(self, color):
        self.borders.override_top_border(color, forced=True)


    ''' recolor bottom border with the specified color
    '''
    def recolor_bottom_border(self, color):
        self.borders.override_bottom_border(color, forced=True)


''' gsheet cell borders object wrapper
'''
class Borders(object):

    ''' constructor
    '''
    def __init__(self, borders_dict=None):
        self.top = None
        self.right = None
        self.bottom = None
        self.left = None

        if borders_dict:
            if 'top' in borders_dict:
                self.top = Border(borders_dict.get('top'))

            if 'right' in borders_dict:
                self.right = Border(borders_dict.get('right'))

            if 'bottom' in borders_dict:
                self.bottom = Border(borders_dict.get('bottom'))

            if 'left' in borders_dict:
                self.left = Border(borders_dict.get('left'))


    ''' string representation
    '''
    def __repr__(self):
        return f"t: [{self.top}], b: [{self.bottom}], l: [{self.left}], r: [{self.right}]"


    ''' override top border with the specified color
    '''
    def override_top_border(self, border_color, forced=False):
        if border_color:
            if self.top is None:
                self.top = Border(None, border_color)
            elif forced:
                self.top.color = border_color


    ''' override bottom border with the specified color
    '''
    def override_bottom_border(self, border_color, forced=False):
        if border_color:
            if self.bottom is None:
                self.bottom = Border(None, border_color)
            elif forced:
                self.bottom.color = border_color


    ''' override left border with the specified color
    '''
    def override_left_border(self, border_color):
        if border_color:
            if self.left is None:
                self.left = Border(None, border_color)


    ''' override right border with the specified color
    '''
    def override_right_border(self, border_color):
        if border_color:
            if self.right is None:
                self.right = Border(None, border_color)


    ''' top border
    '''
    def to_latex_t(self, color_dict):
        t = None
        if self.top:
            t = self.top.to_latex(color_dict)

        return t


    ''' bottom border
    '''
    def to_latex_b(self, color_dict):
        b = None
        if self.bottom:
            b = self.bottom.to_latex(color_dict)

        return b


    ''' left border
    '''
    def to_latex_l(self, color_dict):
        l = None
        if self.left:
            l = self.left.to_latex(color_dict)

        return l


    ''' right border
    '''
    def to_latex_r(self, color_dict):
        r = None
        if self.right:
            r = self.right.to_latex(color_dict)

        return r


''' gsheet cell border object wrapper
'''
class Border(object):

    ''' constructor
    '''
    def __init__(self, border_dict):
        self.style = None
        self.width = None
        self.color = None

        if border_dict:
            self.style = border_dict.get('style')
            self.width = int(border_dict.get('width')) * 0.4
            self.color = RgbColor(border_dict.get('color'))

            # TODO: handle double
            self.style = GSHEET_LATEX_BORDER_MAPPING.get(self.style, 'solid')


    ''' string representation
    '''
    def __repr__(self):
        return f"style: {self.style}, width: {self.width}, color: {self.color}"


    ''' border
    '''
    def to_latex(self, color_dict):
        color_dict[self.color.key()] = self.color.value()
        latex = f"fg={self.color.key()},wd={self.width}pt,dash={self.style}"
        return latex


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
        self.red = 0
        self.green = 0
        self.blue = 0

        if rgb_dict:
            self.red = int(float(rgb_dict.get('red', 0)) * 255)
            self.green = int(float(rgb_dict.get('green', 0)) * 255)
            self.blue = int(float(rgb_dict.get('blue', 0)) * 255)


    ''' string representation
    '''
    def __repr__(self):
        return ''.join('{:02x}'.format(a) for a in [self.red, self.green, self.blue])


    ''' color key for tabularray color name
    '''
    def key(self):
        return ''.join('{:02x}'.format(a) for a in [self.red, self.green, self.blue])


    ''' color value for tabularray color
    '''
    def value(self):
        return ''.join('{:02x}'.format(a) for a in [self.red, self.green, self.blue])


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


''' gsheet text format run object wrapper
'''
class TextFormatRun(object):

    ''' constructor
    '''
    def __init__(self, run_dict=None, default_format=None):
        if run_dict:
            self.start_index = int(run_dict.get('startIndex', 0))
            format = run_dict.get('format')
            new_format = {**default_format, **format}
            self.format = TextFormat(new_format)
        else:
            self.start_index = None
            self.format = None


    ''' generates the latex code
    '''
    def to_latex(self, text, color_dict):
        latex = self.format.to_latex(tex_escape(text[self.start_index:]), color_dict)

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
        self.page_number = False

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
            self.page_number = note_dict.get('page-number') is not None


''' gsheet vertical alignment object wrapper
'''
class VerticalAlignment(object):

    ''' constructor
    '''
    def __init__(self, valign=None):
        if valign:
            self.valign = TBLR_VALIGN.get(valign, 'p')
        else:
            self.valign = TBLR_VALIGN.get('TOP')


''' gsheet horizontal alignment object wrapper
'''
class HorizontalAlignment(object):

    ''' constructor
    '''
    def __init__(self, halign=None):
        if halign:
            self.halign = TBLR_HALIGN.get(halign, 'LEFT')
        else:
            self.halign = TBLR_HALIGN.get('LEFT')


''' Helper for cell span specification
'''
class MultiSpan(object):
    No = 'No'
    FirstCell = 'FirstCell'
    InnerCell = 'InnerCell'
    LastCell = 'LastCell'
