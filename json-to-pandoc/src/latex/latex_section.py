#!/usr/bin/env python3

import json
from pprint import pprint

from pandoc.pandoc_util import *
from latex.latex_cell import *
from latex.latex_util import *

#   ----------------------------------------------------------------------------------------------------------------
#   latex section objects wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' Latex section base object
'''
class LatexSectionBase(object):

    ''' constructor
    '''
    def __init__(self, section_data, config, last_section_was_landscape):
        self._config = config
        self._section_data = section_data

        self.section = self._section_data['section']
        self.heading = section_data['heading']
        self.level = self._section_data['level']
        self.landscape = self._section_data['landscape']
        self.page_numbering = self._section_data['hide-pageno']
        self.section_index = self._section_data['section-index']
        self.section_width = self._section_data['width']
        self.section_break = self._section_data['section-break']
        self.page_break = self._section_data['page-break']
        self.hide_heading = self._section_data['hide-heading']
        self.hide_pageno = self._section_data['hide-pageno']

        self.page_spec = self._config['page-specs']['page-spec'][self._section_data['page-spec']]
        self.margin_spec = self._config['page-specs']['margin-spec'][self._section_data['margin-spec']]

        if self.landscape:
            if last_section_was_landscape:
                self.orientation_changed = False
            else:
                self.orientation_changed = True
        else:
            if last_section_was_landscape:
                self.orientation_changed = True
            else:
                self.orientation_changed = False

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
        header_lines = []
        # section header/footer is only applicable for first section and on section-breaks
        if not self.section_break:
            if self.section_index != 0:
                if self.page_break:
                    header_lines.append(f"\t\\pagebreak")

        else:
            # generate the header block
            header_block = LatexSectionHeader(self.page_spec, self.margin_spec, self.level, self.section, self.heading, self.hide_heading, self.landscape, self.orientation_changed)
            header_lines = header_lines + header_block.to_latex(color_dict)

            # get headers and footers
            if self.header_first.has_content:
                header_lines = header_lines +  list(map(lambda x: f"\t\{x}", self.header_first.to_latex(color_dict)))

            if self.header_odd.has_content:
                header_lines = header_lines +  list(map(lambda x: f"\t{x}", self.header_odd.to_latex(color_dict)))

            if self.header_even.has_content:
                header_lines = header_lines +  list(map(lambda x: f"\t{x}", self.header_even.to_latex(color_dict)))

            if self.footer_first.has_content:
                header_lines = header_lines +  list(map(lambda x: f"\t{x}", self.footer_first.to_latex(color_dict)))

            if self.footer_odd.has_content:
                header_lines = header_lines +  list(map(lambda x: f"\t{x}", self.footer_odd.to_latex(color_dict)))

            if self.footer_even.has_content:
                header_lines = header_lines +  list(map(lambda x: f"\t{x}", self.footer_even.to_latex(color_dict)))

            # now the pagestyle
            header_lines.append(f"\t\\fancypagestyle{{{self.page_style_name}}}{{")
            header_lines.append(f"\t\t\\fancyhf{{}}")
            header_lines.append(f"\t\t\\renewcommand{{\\headrulewidth}}{{0pt}}")
            header_lines.append(f"\t\t\\renewcommand{{\\footrulewidth}}{{0pt}}")
            if self.header_odd.has_content:
                header_lines.append(f"\t\t\t\\fancyhead[O]{{\\{self.header_odd.id}}}")

            if self.header_even.has_content:
                header_lines.append(f"\t\t\t\\fancyhead[E]{{\\{self.header_even.id}}}")

            if self.footer_odd.has_content:
                header_lines.append(f"\t\t\t\\fancyfoot[O]{{\\{self.footer_odd.id}}}")

            if self.footer_even.has_content:
                header_lines.append(f"\t\t\t\\fancyfoot[E]{{\\{self.footer_even.id}}}")

            header_lines.append(f"\t\t}}")
            header_lines.append(f"\t\\pagestyle{{{self.page_style_name}}}")

            # TODO
            header_lines = [f"% PageStyle - [{self.page_style_name}]"] + header_lines

        header_lines = mark_as_latex(header_lines)

        # heading
        heading_lines = []
        if not self.hide_heading:
            hashes = '#' * self.level
            if self.section != '':
                heading_text = f"{hashes} {self.section} {self.heading}".strip()
            else:
                heading_text = f"{hashes} {self.heading}".strip()

            # headings are styles based on level
            if self.level != 0:
                heading_lines.append(heading_text)
                heading_lines.append('\n')
            else:
                heading_lines.append(f"\\titlestyle{{{heading_text}}}")
                heading_lines = mark_as_latex(heading_lines)

        return header_lines + heading_lines



''' Latex table section object
'''
class LatexTableSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config, last_section_was_landscape):
        super().__init__(section_data, config, last_section_was_landscape)


    ''' generates the latex code
    '''
    def to_latex(self, color_dict):
        section_lines = super().to_latex(color_dict)

        # get the contents
        content_lines = self.section_contents.to_latex(color_dict)

        return section_lines + content_lines



''' Latex ToC section object
'''
class LatexToCSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config, last_section_was_landscape):
        super().__init__(section_data, config, last_section_was_landscape)


    ''' generates the latex code
    '''
    def to_latex(self, color_dict):
        section_lines = super().to_latex(color_dict)

        # table-of-contents
        content_lines = []
        content_lines.append("\\renewcommand{\\contentsname}{}")
        content_lines.append("\\vspace{-0.5in}")
        content_lines.append("\\tableofcontents")
        content_lines.append("\\addtocontents{toc}{~\\hfill\\textbf{Page}\\par}")
        content_lines = mark_as_latex(content_lines)

        return section_lines + content_lines



''' Latex LoT section object
'''
class LatexLoTSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config, last_section_was_landscape):
        super().__init__(section_data, config, last_section_was_landscape)


    ''' generates the latex code
    '''
    def to_latex(self, color_dict):
        section_lines = super().to_latex(color_dict)

        # table-of-contents
        content_lines = []
        content_lines.append("\\renewcommand{\\listtablename}{}")
        content_lines.append("\\vspace{-0.4in}")
        content_lines.append("\\listoftables")
        content_lines.append("\\addtocontents{lot}{~\\hfill\\textbf{Page}\\par}")
        content_lines = mark_as_latex(content_lines)

        return section_lines + content_lines



''' Latex LoF section object
'''
class LatexLoFSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config, last_section_was_landscape):
        super().__init__(section_data, config, last_section_was_landscape)


    ''' generates the latex code
    '''
    def to_latex(self, color_dict):
        section_lines = super().to_latex(color_dict)

        # table-of-contents
        content_lines = []
        content_lines.append("\\renewcommand{\\listfigurename}{}")
        content_lines.append("\\vspace{-0.4in}")
        content_lines.append("\\listoffigures")
        content_lines.append("\\addtocontents{lof}{~\\hfill\\textbf{Page}\\par}")
        content_lines = mark_as_latex(content_lines)

        return section_lines + content_lines



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
            # self.default_format = CellFormat(properties.get('defaultFormat'))

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
                        self.cell_matrix.append(Row(r, row_data, self.default_format, self.section_width, self.column_widths, self.row_metadata_list[r].inches))
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

            # get the total height of the first_cell when merged
            for r in range(first_row + 1, last_row):
                first_cell.cell_height = first_cell.cell_height + self.cell_matrix[r].row_height + ROWSEP * 2

            if col_span > 1:
                first_cell.merge_spec.multi_col = MultiSpan.FirstCell
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
                row_height = next_row_object.row_height

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
                        next_cell_in_row = Cell(r, c, None, first_cell.default_format, first_cell.column_widths, row_height)
                        next_row_object.insert_cell(c, next_cell_in_row)

                    if next_cell_in_row.is_empty:
                        # debug(f"..cell [{r+1},{c+1}] is empty")
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
        last_block_is_a_paragraph = False
        for block in self.content_list:
            # if current block is a paragraph and last block was also a paragraph, insert a newline
            if isinstance(block, LatexParagraph) and last_block_is_a_paragraph:
                latex_lines.append(f"\t\\\\[{10}pt]")

            latex_lines = latex_lines + block.to_latex(longtable=True, color_dict=color_dict)

            # keep track of the block as the previous block
            if isinstance(block, LatexParagraph):
                last_block_is_a_paragraph = True
            else:
                last_block_is_a_paragraph = False

        latex_lines = mark_as_latex(latex_lines)

        return latex_lines


'''
'''
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
            block_lines = block.to_latex(longtable=False, color_dict=color_dict, strip_comments=True, header_footer=self.header_footer)
            # block_lines = list(map(lambda x: f"\t{x}", block_lines))
            latex_lines = latex_lines + block_lines

            first_block = False

        latex_lines.append(f"\t}}")

        return latex_lines


''' Latex Block object wrapper base class (plain latex, table, header etc.)
'''
class LatexBlock(object):

    ''' generates latex code
    '''
    def to_latex(self, longtable, color_dict, strip_comments=False, header_footer=None):
        pass


''' Latex Header Block wrapper
'''
class LatexSectionHeader(LatexBlock):

    ''' constructor
    '''
    def __init__(self, page_spec, margin_spec, level, section, heading, hide_heading, landscape, orientation_changed):
        self.page_spec =page_spec
        self.margin_spec = margin_spec
        self.level = level
        self.section = section
        self.heading = heading
        self.hide_heading = hide_heading
        self.landscape = landscape
        self.orientation_changed = orientation_changed


    ''' generates latex code
    '''
    def to_latex(self, color_dict, strip_comments=False, header_footer=None):
        header_lines = []
        paper = self.page_spec['name']
        page_width = self.page_spec['width']
        page_height = self.page_spec['height']
        top_margin = self.margin_spec['top']
        bottom_margin = self.margin_spec['bottom']
        left_margin = self.margin_spec['left']
        right_marhin = self.margin_spec['right']

        header_lines.append(f"\t\\pagebreak")

        if self.orientation_changed:
            if self.landscape:
                header_lines.append(f"\t\\newgeometry{{{paper}, top={top_margin}in, bottom={bottom_margin}in, left={left_margin}in, right={right_marhin}in, landscape}}")
            else:
                header_lines.append(f"\t\\restoregeometry")

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
    def to_latex(self, longtable, color_dict, strip_comments=False, header_footer=None):
        # table template
        template_lines = []
        if longtable:
            template_lines.append(f"\\DefTblrTemplate{{caption-tag}}{{default}}{{}}")
            template_lines.append(f"\\DefTblrTemplate{{caption-sep}}{{default}}{{}}")
            template_lines.append(f"\\DefTblrTemplate{{caption-text}}{{default}}{{}}")
            template_lines.append(f"\\DefTblrTemplate{{conthead}}{{default}}{{}}")
            template_lines.append(f"\\DefTblrTemplate{{contfoot}}{{default}}{{}}")
            template_lines.append(f"\\DefTblrTemplate{{conthead-text}}{{default}}{{}}")
            template_lines.append(f"\\DefTblrTemplate{{contfoot-text}}{{default}}{{}}")
            table_type = "longtblr"
        else:
            table_type = "tblr"

        # optional inner keys
        table_inner_keys = [f"caption=,", f"entry=none,", f"label=none,", f"presep={0}pt,", f"postsep={0}pt,"]
        table_inner_keys = list(map(lambda x: f"\t\t{x}", table_inner_keys))

        # table spec keys
        table_col_spec = ''.join([f"p{{{i}in}}" for i in self.column_widths])
        table_col_spec = f"colspec={{{table_col_spec}}},"
        table_rulesep = f"rulesep={0}pt,"
        table_stretch = f"stretch={1.0},"
        table_vspan = f"vspan=even,"
        table_hspan = f"hspan=minimal,"
        table_rows = f"rows={{ht={12}pt}},"
        table_rowhead = f"rowhead={self.header_row_count},"

        # table_spec_keys = [table_col_spec, table_rulesep, table_stretch, table_vspan, table_hspan, table_rows]
        table_spec_keys = [table_col_spec, table_rulesep, table_stretch, table_vspan, table_hspan]
        if longtable:
            table_spec_keys.append(table_rowhead)

        table_spec_keys = list(map(lambda x: f"\t\t{x}", table_spec_keys))

        table_lines = []
        table_lines = table_lines + template_lines
        table_lines.append(f"\\begin{{{table_type}}}[")
        table_lines = table_lines + table_inner_keys
        table_lines.append("\t]{")
        table_lines = table_lines + table_spec_keys


        # generate row formats
        r = 1
        for row in self.table_cell_matrix:
            row_format_line = f"\t\t{row.row_format_latex(r, color_dict)}"
            table_lines.append(row_format_line)
            r = r + 1

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
        table_lines.append(f"\t}}")

        # generate cell values
        for row in self.table_cell_matrix:
            row_lines = list(map(lambda x: f"\t{x}", row.cell_content_latex(include_formatting=False, color_dict=color_dict, strip_comments=strip_comments, header_footer=header_footer)))
            table_lines = table_lines + row_lines

        table_lines.append(f"\t\\end{{{table_type}}}")
        table_lines = list(map(lambda x: f"\t{x}", table_lines))

        if not strip_comments:
            table_lines = [f"% LatexTable: ({self.start_row+1}-{self.end_row+1}) : {self.row_count} rows"] + table_lines

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
    def to_latex(self, longtable, color_dict, strip_comments=False):
        block_lines = []
        if not strip_comments:
            block_lines.append(f"% LatexParagraph: row {self.row_number+1}")

        # TODO 3: generate the block
        if len(self.data_row.cells) > 0:
            row_text = self.data_row.get_cell(0).content_latex(include_formatting=True, color_dict=color_dict)
            row_lines = list(map(lambda x: f"\t{x}", row_text))
            block_lines = block_lines + row_lines

        return block_lines
