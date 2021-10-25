#!/usr/bin/env python3

import json
from helper.pandoc.pandoc_util import *
from helper.latex.latex_cell import *
from helper.latex.latex_util import *

#   ----------------------------------------------------------------------------------------------------------------
#   latex section objects wrappers
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

        # the colors dict to be returned
        self.color_dict = {}

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
        pass


''' Latex section object
'''
class LatexSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, section_spec):
        super().__init__(section_data, section_spec)


    ''' processes the cells to
        1. identify missing cells/contents for merged cells
        2. generate the proper order of tables and blocks
    '''
    def process(self):

        # first we identify the missing cells or blank cells for merged spans
        for merge in self.merge_list:
            first_row = merge.start_row
            first_col = merge.start_col
            last_row = merge.end_row
            last_col = merge.end_col
            row_span = merge.row_span
            col_span = merge.col_span
            first_row_object = self.cell_matrix[first_row]
            first_cell = first_row_object.get_cell(first_col)

            if first_cell is None:
                warn(f"cell [{first_row},{first_col}] starts a span, but it is not there")
                continue

            # get the total width of the first_cell when merged
            for c in range(first_col + 1, last_col):
                first_cell.cell_width = first_cell.cell_width + first_cell.column_widths[c] + COLSEP * 2

            if col_span > 1:
                first_cell.merge_spec.multi_col = MultiSpan.FirstCell
                first_cell.cell_width = first_cell.cell_width + ((col_span - 1) * 2 * COLSEP)
            else:
                first_cell.merge_spec.multi_col = MultiSpan.No

            if row_span > 1:
                first_cell.merge_spec.multi_row = MultiSpan.FirstCell
            else:
                first_cell.merge_spec.multi_row = MultiSpan.No

            first_cell.merge_spec.col_span = col_span
            first_cell.merge_spec.row_span = row_span

            # for row spans, subsequent cells in the same column of the FirstCell will be either empty or missing, iterate through the next rows
            for r in range(first_row, last_row):
                next_row_object = self.cell_matrix[r]

                # we may have empty cells in this same row which are part of this column merge, we need to mark their multi_col property correctly
                for c in range(first_col, last_col):
                    # exclude the very first cell
                    if r == first_row and c == first_col:
                        continue

                    # debug(f"..cell [{r+1},{c+1}] is part of column merge")
                    next_cell_in_row = next_row_object.get_cell(c)

                    if next_cell_in_row is None:
                        # the cell may not be existing at all, we have to create
                        # debug(f"..cell [{r+1},{c+1}] does not exist, to be inserted")
                        next_cell_in_row = Cell(r, c, None, first_cell.default_format, first_cell.column_widths)
                        next_row_object.insert_cell(c, next_cell_in_row)

                    if next_cell_in_row.is_empty:
                        # the cell is a newly inserted one, its format should be the same (for borders, colors) as the first cell so that we can draw borders properly
                        next_cell_in_row.copy_format_from(first_cell)

                        # mark cells for multicol only if it is multicol
                        if col_span > 1:
                            if c == first_col:
                                # the last cell of the merge to be marked as LastCell
                                # debug(f"..cell [{r+1},{c+1}] is the LastCell of the column merge")
                                next_cell_in_row.mark_multicol(MultiSpan.FirstCell)

                            elif c == last_col-1:
                                # the last cell of the merge to be marked as LastCell
                                # debug(f"..cell [{r+1},{c+1}] is the LastCell of the column merge")
                                next_cell_in_row.mark_multicol(MultiSpan.LastCell)

                            else:
                                # the inner cells of the merge to be marked as InnerCell
                                # debug(f"..cell [{r+1},{c+1}] is an InnerCell of the column merge")
                                next_cell_in_row.mark_multicol(MultiSpan.InnerCell)

                        else:
                            next_cell_in_row.mark_multicol(MultiSpan.No)


                        # mark cells for multirow only if it is multirow
                        if row_span > 1:
                            if r == first_row:
                                # the last cell of the merge to be marked as LastCell
                                # debug(f"..cell [{r+1},{c+1}] is the LastCell of the row merge")
                                next_cell_in_row.mark_multirow(MultiSpan.FirstCell)

                            elif r == last_row-1:
                                # the last cell of the merge to be marked as LastCell
                                # debug(f"..cell [{r+1},{c+1}] is the LastCell of the row merge")
                                next_cell_in_row.mark_multirow(MultiSpan.LastCell)

                            else:
                                # the inner cells of the merge to be marked as InnerCell
                                # debug(f"..cell [{r+1},{c+1}] is an InnerCell of the row merge")
                                next_cell_in_row.mark_multirow(MultiSpan.InnerCell)

                        else:
                            next_cell_in_row.mark_multirow(MultiSpan.No)


                    else:
                        warn(f"..cell [{r+1},{c+1}] is not empty, it must be part of another column/row merge which is an issue")


        # let us see how the cells look now
        # for row in self.cell_matrix:
        #     debug(row.row_name)
        #     for cell in row.cells:
        #         debug(f".. {cell}")

        return


    ''' processes the cells to split the cells into tables and blocks nd ordering the tables and blocks properly
    '''
    def split(self):
        # we have a concept of in-cell content and out-of-cell content
        # in-cell contents are treated as part of a table, while out-of-cell contents are treated as independent paragraphs, images etc. (blocks)
        next_table_starts_in_row = 0
        next_table_ends_in_row = 0
        for r in range(0, self.row_count):
            # the first cell of the row tells us whether it is in-cell or out-of-cell
            data_row = self.cell_matrix[r]
            if data_row.is_out_of_table():
                # there may be a pending/running table
                if r > next_table_starts_in_row:
                    table = LatexTable(self.cell_matrix, next_table_starts_in_row, r - 1, self.column_widths)
                    self.content_list.append(table)

                block = LatexParagraph(data_row, r)
                self.content_list.append(block)

                next_table_starts_in_row = r + 1

            # the row may start with a note of repeat-rows which means that a new table is atarting
            elif data_row.is_table_start():
                # there may be a pending/running table
                if r > next_table_starts_in_row:
                    table = LatexTable(self.cell_matrix, next_table_starts_in_row, r - 1, self.column_widths)
                    self.content_list.append(table)

                    next_table_starts_in_row = r

            else:
                next_table_ends_in_row = r

        # there may be a pending/running table
        if next_table_ends_in_row >= next_table_starts_in_row:
            table = LatexTable(self.cell_matrix, next_table_starts_in_row, next_table_ends_in_row, self.column_widths)
            self.content_list.append(table)


    ''' generates the latex code
    '''
    def to_latex(self):
        # process and split
        self.process()
        self.split()

        latex_lines = []

        # iterate to through tables and blocks contents
        for block in self.content_list:
            latex_lines = latex_lines + block.to_latex(self.color_dict)

        return latex_lines, self.color_dict


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
            latex_lines = latex_lines + block.to_latex(self.color_dict)

        return latex_lines, self.color_dict


''' Latex Block object wrapper base class (plain latex, table, header etc.)
'''
class LatexBlock(object):

    ''' generates latex code
    '''
    def to_latex(self, color_dict):
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
    def to_latex(self, color_dict):
        header_lines = []
        header_lines.append('')
        header_lines.append(begin_latex())
        header_lines.append(f"% LatexSection: {self.title}")

        if self.section_break.startswith('newpage_'):
            header_lines.append(f"\t\\newpage")

        header_lines.append(f"\t\\pdfpagewidth {self.section_spec['page_width']}in")
        header_lines.append(f"\t\\pdfpageheight {self.section_spec['page_height']}in")
        header_lines.append(f"\t\\newgeometry{{top={self.section_spec['top_margin']}in, bottom={self.section_spec['bottom_margin']}in, left={self.section_spec['left_margin']}in, right={self.section_spec['right_margin']}in}}")

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
        self.table_name = f"LatexTable: {self.start_row+1}-{self.end_row+1}[{self.row_count}]"

        # header row if any
        self.header_row_count = self.table_cell_matrix[0].get_cell(0).note.header_rows


    ''' generates the latex code
    '''
    def to_latex(self, color_dict):
        table_col_spec = ''.join([f"p{{{i}in}}" for i in self.column_widths])
        table_col_spec = f"colspec={{{table_col_spec}}}"
        table_rulesep = f"rulesep={0}pt"
        table_stretch = f"stretch={1.0}"
        table_vspan = f"vspan=default"
        table_hspan = f"hspan=default"
        table_rows = f"rows={{ht={12}pt}}"
        table_rowhead = f"rowhead={self.header_row_count}"

        table_lines = []

        table_lines.append(begin_latex())
        table_lines.append(f"% LatexTable: ({self.start_row+1}-{self.end_row+1}) : {self.row_count} rows")
        table_lines.append(f"\t\\DefTblrTemplate{{caption-tag}}{{default}}{{}}")
        table_lines.append(f"\t\\DefTblrTemplate{{caption-sep}}{{default}}{{}}")
        table_lines.append(f"\t\\DefTblrTemplate{{caption-text}}{{default}}{{}}")
        table_lines.append(f"\t\\DefTblrTemplate{{conthead-text}}{{default}}{{}}")
        table_lines.append(f"\t\\DefTblrTemplate{{contfoot-text}}{{default}}{{}}")

        table_lines.append(f"\t\\begin{{longtblr}}[l]{{\n\t\t{table_col_spec},\n\t\t{table_rulesep},\n\t\t{table_stretch},\n\t\t{table_vspan},\n\t\t{table_hspan},\n\t\t{table_rows},\n\t\t{table_rowhead},")

        # generate cell formats
        r = 1
        for row in self.table_cell_matrix:
            cell_format_lines = list(map(lambda x: f"\t\t{x}", row.cell_format_latex(r, color_dict)))
            table_lines = table_lines + cell_format_lines
            r = r + 1

        # generate vertical borders
        r = 1
        for row in self.table_cell_matrix:
            v_lines = list(map(lambda x: f"\t\t{x}", row.vertical_borders_latex(r, color_dict)))
            table_lines = table_lines + v_lines
            r = r + 1

        # close the table definition
        table_lines.append(f"\t}}\n")

        # generate cell values
        for row in self.table_cell_matrix:
            row_lines = list(map(lambda x: f"\t{x}", row.cell_content_latex(color_dict)))
            table_lines = table_lines + row_lines

        table_lines.append(f"\t\\end{{longtblr}}")
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
    def to_latex(self, color_dict):
        block_lines = []
        block_lines.append(begin_latex())
        block_lines.append(f"% LatexParagraph: row {self.row_number+1}")

        # TODO 3: generate the block
        if len(self.data_row.cells) > 0:
            row_text = self.data_row.get_cell(0).content_latex(color_dict)
            row_lines = list(map(lambda x: f"\t{x}", row_text))
            block_lines = block_lines + row_lines

        block_lines.append(end_latex())
        return block_lines
