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
    def __init__(self, section_data, section_spec, section_index):
        self.section_spec = section_spec
        self.section_break = section_data['section-break']
        self.section_width = float(self.section_spec['page_width']) - float(self.section_spec['left_margin']) - float(self.section_spec['right_margin']) - float(self.section_spec['gutter'])

        self.no_heading = section_data['no-heading']
        self.section = section_data['section']
        self.heading = section_data['heading']
        self.level = section_data['level']
        self.page_numbering = section_data['page-number']
        self.title = f"{self.section} : {self.heading}"
        self.section_index = section_index
        self.page_style_name = f"pagestyle{self.section_index}"

        # headers and footers
        self.header_first = LatexPageHeaderFooter(section_data['header-first'], self.section_width, self.section_index, header_footer='header', odd_even='first')
        self.header_odd = LatexPageHeaderFooter(section_data['header-odd'], self.section_width, self.section_index, header_footer='header', odd_even='odd')
        self.header_even = LatexPageHeaderFooter(section_data['header-even'], self.section_width, self.section_index, header_footer='header', odd_even='even')
        self.footer_first = LatexPageHeaderFooter(section_data['footer-first'], self.section_width, self.section_index, header_footer='footer', odd_even='first')
        self.footer_odd = LatexPageHeaderFooter(section_data['footer-odd'], self.section_width, self.section_index, header_footer='footer', odd_even='odd')
        self.footer_even = LatexPageHeaderFooter(section_data['footer-even'], self.section_width, self.section_index, header_footer='footer', odd_even='even')

        self.section_contents = LatexContent(section_data.get('contents'), self.section_width, self.section_index)


    ''' generates the latex code
    '''
    def to_latex(self, color_dict):
        header_footer_lines = []

        # get headers and footers
        if self.header_first.has_content:
            header_footer_lines = header_footer_lines + self.header_first.to_latex(color_dict)

        if self.header_odd.has_content:
            header_footer_lines = header_footer_lines + self.header_odd.to_latex(color_dict)

        if self.header_even.has_content:
            header_footer_lines = header_footer_lines + self.header_even.to_latex(color_dict)

        if self.footer_first.has_content:
            header_footer_lines = header_footer_lines + self.footer_first.to_latex(color_dict)

        if self.footer_odd.has_content:
            header_footer_lines = header_footer_lines + self.footer_odd.to_latex(color_dict)

        if self.footer_even.has_content:
            header_footer_lines = header_footer_lines + self.footer_even.to_latex(color_dict)

        header_footer_lines = mark_as_latex(header_footer_lines)

        # now the pagestyle
        pagestyle_lines = []
        pagestyle_lines.append(f"\\fancypagestyle{{{self.page_style_name}}}{{")
        pagestyle_lines.append(f"\t\\fancyhf{{}}")
        pagestyle_lines.append(f"\t\\renewcommand{{\headrulewidth}}{{0pt}}")
        if self.header_odd.has_content:
            pagestyle_lines.append(f"\t\\fancyhead[O]{{\\{self.header_odd.id}}}")

        if self.header_even.has_content:
            pagestyle_lines.append(f"\t\\fancyhead[E]{{\\{self.header_even.id}}}")

        if self.footer_odd.has_content:
            pagestyle_lines.append(f"\t\\fancyfoot[O]{{\\{self.footer_odd.id}}}")

        if self.footer_even.has_content:
            pagestyle_lines.append(f"\t\\fancyfoot[E]{{\\{self.footer_even.id}}}")

        pagestyle_lines.append(f"}}")
        pagestyle_lines = mark_as_latex(pagestyle_lines)

        # generate the header block
        header_block = LatexSectionHeader(self.section_spec, self.section_break, self.level, self.no_heading, self.section, self.heading, self.title, self.page_style_name)
        header_lines = header_block.to_latex(color_dict)

        # get the contents
        content_lines = self.section_contents.to_latex(color_dict)

        return header_footer_lines + pagestyle_lines + header_lines + content_lines


''' Latex table section object
'''
class LatexTableSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, section_spec, section_index):
        super().__init__(section_data, section_spec, section_index)


    ''' generates the latex code
    '''
    def to_latex(self, color_dict):
        return super().to_latex(color_dict)


''' Latex section object
'''
class LatexToCSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, section_spec, section_index):
        super().__init__(section_data, section_spec, section_index)


    ''' generates the latex code
    '''
    def to_latex(self, color_dict):
        return super().to_latex(color_dict)


''' Latex section content base object
'''
class LatexContent(object):

    ''' constructor
    '''
    def __init__(self, content_data, section_width, section_index):
        self.section_width = section_width
        self.section_index = section_index
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

        # we need a list to hold the tables and block for the cells
        self.content_list = []

        if content_data:
            self.has_content = True

            properties = content_data.get('properties')
            self.default_format = CellFormat(properties.get('defaultFormat'))

            sheets = content_data.get('sheets')
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

            # process and split
            self.process()
            self.split()

        else:
            self.has_content = False


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
    def to_latex(self, color_dict):
        latex_lines = []

        # iterate through tables and blocks contents
        first_block = True
        for block in self.content_list:
            if not first_block:
                latex_lines.append('')

            latex_lines = latex_lines + block.to_latex(longtable=True, color_dict=color_dict)
            first_block = False

        latex_lines = mark_as_latex(latex_lines)

        return latex_lines


class LatexPageHeaderFooter(LatexContent):

    ''' constructor
        header_footer : header/footer
        odd_even      : first/odd/even
    '''
    def __init__(self, content_data, section_width, section_index, header_footer, odd_even):
        super().__init__(content_data, section_width, section_index)
        self.header_footer, self.odd_even = header_footer, odd_even
        self.id = f"{self.header_footer}{self.odd_even}{section_index}"


    ''' generates the latex code
    '''
    def to_latex(self, color_dict):
        latex_lines = []

        latex_lines.append(f"\\newcommand{{\\{self.id}}}{{%")

        # iterate through tables and blocks contents
        first_block = True
        for block in self.content_list:
            if not first_block:
                latex_lines.append('')

            latex_lines = latex_lines + block.to_latex(longtable=False, color_dict=color_dict)
            first_block = False

        latex_lines.append(f"}}")

        return latex_lines


''' Latex Block object wrapper base class (plain latex, table, header etc.)
'''
class LatexBlock(object):

    ''' generates latex code
    '''
    def to_latex(self, longtable, color_dict):
        pass


''' Latex Header Block wrapper
'''
class LatexSectionHeader(LatexBlock):

    ''' constructor
    '''
    def __init__(self, section_spec, section_break, level, no_heading, section, heading, title, page_style_name):
        self.section_spec, self.section_break, self.level, self.no_heading, self.section, self.heading, self.title, self.page_style_name = section_spec, section_break, level, no_heading, section, heading, title, page_style_name

    ''' generates latex code
    '''
    def to_latex(self, color_dict):
        header_lines = []
        # header_lines.append('')
        header_lines.append(f"% LatexSection: {self.title}")

        if self.section_break.startswith('newpage_'):
            header_lines.append(f"\t\\newpage")

        header_lines.append(f"\t\\pdfpagewidth {self.section_spec['page_width']}in")
        header_lines.append(f"\t\\pdfpageheight {self.section_spec['page_height']}in")
        header_lines.append(f"\t\\newgeometry{{top={self.section_spec['top_margin']}in, bottom={self.section_spec['bottom_margin']}in, left={self.section_spec['left_margin']}in, right={self.section_spec['right_margin']}in}}")
        header_lines.append(f"\t\\pagestyle{{{self.page_style_name}}}")
        header_lines = mark_as_latex(header_lines)

        # heading
        heading_lines = []
        if not self.no_heading:
            if self.section != '':
                heading_text = f"{'#' * self.level} {self.section} - {self.heading}".strip()
            else:
                heading_text = f"{'#' * self.level} {self.heading}".strip()

            # headings are styles based on level
            if self.level != 0:
                heading_lines.append(heading_text)
                heading_lines.append('\n')
            else:
                heading_lines.append(f"\\titlestyle{{{heading_text}}}")
                heading_lines = mark_as_latex(heading_lines)

        return header_lines + heading_lines


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
    def to_latex(self, longtable, color_dict):
        table_col_spec = ''.join([f"p{{{i}in}}" for i in self.column_widths])
        table_col_spec = f"colspec={{{table_col_spec}}}"
        table_rulesep = f"rulesep={0}pt"
        table_stretch = f"stretch={1.0}"
        table_vspan = f"vspan=default"
        table_hspan = f"hspan=default"
        table_rows = f"rows={{ht={12}pt}}"
        table_rowhead = f"rowhead={self.header_row_count}"

        table_inner_keys = "caption=,entry=none,label=none,presep=0pt,postsep=0pt"


        table_lines = []

        table_lines.append(f"% LatexTable: ({self.start_row+1}-{self.end_row+1}) : {self.row_count} rows")
        if longtable:
            table_lines.append(f"\t\\DefTblrTemplate{{caption-tag}}{{default}}{{}}")
            table_lines.append(f"\t\\DefTblrTemplate{{caption-sep}}{{default}}{{}}")
            table_lines.append(f"\t\\DefTblrTemplate{{caption-text}}{{default}}{{}}")
            table_lines.append(f"\t\\DefTblrTemplate{{conthead-text}}{{default}}{{}}")
            table_lines.append(f"\t\\DefTblrTemplate{{contfoot-text}}{{default}}{{}}")
            table_type = "longtblr"
        else:
            table_type = "tblr"

        table_lines.append(f"\t\\begin{{{table_type}}}[{table_inner_keys}]{{\n\t\t{table_col_spec},\n\t\t{table_rulesep},\n\t\t{table_stretch},\n\t\t{table_vspan},\n\t\t{table_hspan},\n\t\t{table_rows},\n\t\t{table_rowhead},")

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

        table_lines.append(f"\t\\end{{{table_type}}}")

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
    def to_latex(self, longtable, color_dict):
        block_lines = []
        block_lines.append(f"% LatexParagraph: row {self.row_number+1}")

        # TODO 3: generate the block
        if len(self.data_row.cells) > 0:
            row_text = self.data_row.get_cell(0).content_latex(color_dict)
            row_lines = list(map(lambda x: f"\t{x}", row_text))
            block_lines = block_lines + row_lines

        return block_lines
