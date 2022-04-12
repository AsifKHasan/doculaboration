#!/usr/bin/env python3

import json
from odt.odt_cell import *
from odt.odt_util import *

#   ----------------------------------------------------------------------------------------------------------------
#   odt section objects wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' Odt section base object
'''
class OdtSectionBase(object):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        self._config = config
        self._section_data = section_data


    ''' generates the odt code
    '''
    def to_odt(self, odt):
        # Headings
        # identify what style the heading will be and its content
        if not self._section_data['hide-heading']:
            heading_text = self._section_data['heading']
            if self._section_data['section'] != '':
                heading_text = f"{self._section_data['section']} {heading_text}"

            if self._section_data['level'] == 0:
                parent_style_name = 'Title'
            else:
                parent_style_name = f"Heading_20_{self._section_data['level']}"

            # debug(f"..... {style_name} - {heading_text}")
        else:
            heading_text = ''
            parent_style_name = 'Text_20_body'

        # handle section-break and page-break
        if self._section_data['section-break']:
            # if it is a new-section, we create a new paragraph-style based on parent_style_name with the master-page and apply it
            style_name = f"{self._section_data['section-index']}-P0-with-section-break"
            paragraph_style_name = create_paragraph_style(odt, style_name, parent_style_name, page_break=False, master_page_name=self._section_data['master-page'])
            create_paragraph(odt, style_name, heading_text)
        else:
            if self._section_data['page-break']:
                # if it is a new-page, we create a new paragraph-style based on parent_style_name with the page-break and apply it
                style_name = f"{self._section_data['section-index']}-P0-with-page-break"
                create_paragraph_style(odt, style_name, parent_style_name, page_break=True, master_page_name=None)
                create_paragraph(odt, style_name, heading_text)
            else:
                # just write it
                style_name = f"{self._section_data['section-index']}-P0"
                create_paragraph_style(odt, style_name, parent_style_name, page_break=False, master_page_name=None)
                create_paragraph(odt, parent_style_name, heading_text)



''' Odt table section object
'''
class OdtTableSection(OdtSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data, config)


    ''' generates the odt code
    '''
    def to_odt(self, odt):
        super().to_odt(odt)


''' Odt ToC section object
'''
class OdtToCSection(OdtSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data, config)


    ''' generates the odt code
    '''
    def to_odt(self, odt):
        super().to_odt(odt)
        create_toc(odt)


''' Odt LoT section object
'''
class OdtLoTSection(OdtSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data, config)


    ''' generates the odt code
    '''
    def to_odt(self, odt):
        super().to_odt(odt)


''' Odt LoF section object
'''
class OdtLoFSection(OdtSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data, config)


    ''' generates the odt code
    '''
    def to_odt(self, odt):
        super().to_odt(odt)


''' Odt section content base object
'''
class OdtContent(object):

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
                        debug(f"..cell [{r+1},{c+1}] does not exist, to be inserted")
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
                        # pass


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
                    table = OdtTable(self.cell_matrix, next_table_starts_in_row, r - 1, self.column_widths)
                    self.content_list.append(table)

                block = OdtParagraph(data_row, r)
                self.content_list.append(block)

                next_table_starts_in_row = r + 1

            # the row may start with a note of repeat-rows which means that a new table is atarting
            elif data_row.is_table_start():
                # there may be a pending/running table
                if r > next_table_starts_in_row:
                    table = OdtTable(self.cell_matrix, next_table_starts_in_row, r - 1, self.column_widths)
                    self.content_list.append(table)

                    next_table_starts_in_row = r

            else:
                next_table_ends_in_row = r

        # there may be a pending/running table
        if next_table_ends_in_row >= next_table_starts_in_row:
            table = OdtTable(self.cell_matrix, next_table_starts_in_row, next_table_ends_in_row, self.column_widths)
            self.content_list.append(table)


    ''' generates the odt code
    '''
    def to_odt(self, color_dict):
        odt_lines = []

        # iterate through tables and blocks contents
        last_block_is_a_paragraph = False
        for block in self.content_list:
            # if current block is a paragraph and last block was also a paragraph, insert a newline
            if isinstance(block, OdtParagraph) and last_block_is_a_paragraph:
                odt_lines.append(f"\t\\\\[{10}pt]")

            odt_lines = odt_lines + block.to_odt(longtable=True, color_dict=color_dict)

            # keep track of the block as the previous block
            if isinstance(block, OdtParagraph):
                last_block_is_a_paragraph = True
            else:
                last_block_is_a_paragraph = False

        odt_lines = mark_as_odt(odt_lines)

        return odt_lines


class OdtPageHeaderFooter(OdtContent):

    ''' constructor
        header_footer : header/footer
        odd_even      : first/odd/even
    '''
    def __init__(self, content_data, section_width, section_index, header_footer, odd_even):
        super().__init__(content_data, section_width, section_index)
        self.header_footer, self.odd_even = header_footer, odd_even
        self.id = f"{self.header_footer}{self.odd_even}{section_index}"


    ''' generates the odt code
    '''
    def to_odt(self, color_dict):
        odt_lines = []

        odt_lines.append(f"\\newcommand{{\\{self.id}}}{{%")

        # iterate through tables and blocks contents
        first_block = True
        for block in self.content_list:
            block_lines = block.to_odt(longtable=False, color_dict=color_dict, strip_comments=True, header_footer=self.header_footer)
            # block_lines = list(map(lambda x: f"\t{x}", block_lines))
            odt_lines = odt_lines + block_lines

            first_block = False

        odt_lines.append(f"\t}}")

        return odt_lines


''' Odt Block object wrapper base class (plain odt, table, header etc.)
'''
class OdtBlock(object):

    ''' generates odt code
    '''
    def to_odt(self, longtable, color_dict, strip_comments=False, header_footer=None):
        pass


''' Odt Header Block wrapper
'''
class OdtSectionHeader(OdtBlock):

    ''' constructor
    '''
    def __init__(self, section_spec, level, no_heading, section, heading, newpage, landscape, orientation_changed):
        self.section_spec, self.level, self.no_heading, self.section, self.heading, self.newpage, self.landscape, self.orientation_changed = section_spec, level, no_heading, section, heading, newpage, landscape, orientation_changed

    ''' generates odt code
    '''
    def to_odt(self, color_dict, strip_comments=False, header_footer=None):
        header_lines = []
        paper = 'a4paper'
        page_width = self.section_spec['page_width']
        page_height = self.section_spec['page_height']
        top_margin = self.section_spec['top_margin']
        bottom_margin = self.section_spec['bottom_margin']
        left_margin = self.section_spec['left_margin']
        right_marhin = self.section_spec['right_margin']

        if self.newpage:
            header_lines.append(f"\t\\pagebreak")

        if self.orientation_changed:
            if self.landscape:
                header_lines.append(f"\t\\newgeometry{{{paper}, top={top_margin}in, bottom={bottom_margin}in, left={left_margin}in, right={right_marhin}in, landscape}}")
            else:
                header_lines.append(f"\t\\restoregeometry")

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
                heading_lines = mark_as_odt(heading_lines)

        return header_lines, heading_lines


''' Odt Table object wrapper
'''
class OdtTable(OdtBlock):

    ''' constructor
    '''
    def __init__(self, cell_matrix, start_row, end_row, column_widths):
        self.start_row, self.end_row, self.column_widths = start_row, end_row, column_widths
        self.table_cell_matrix = cell_matrix[start_row:end_row+1]
        self.row_count = len(self.table_cell_matrix)
        self.table_name = f"OdtTable: {self.start_row+1}-{self.end_row+1}[{self.row_count}]"

        # header row if any
        self.header_row_count = self.table_cell_matrix[0].get_cell(0).note.header_rows


    ''' generates the odt code
    '''
    def to_odt(self, longtable, color_dict, strip_comments=False, header_footer=None):
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
            row_format_line = f"\t\t{row.row_format_odt(r, color_dict)}"
            table_lines.append(row_format_line)
            r = r + 1

        # generate cell formats
        r = 1
        for row in self.table_cell_matrix:
            cell_format_lines = list(map(lambda x: f"\t\t{x}", row.cell_format_odt(r, color_dict)))
            table_lines = table_lines + cell_format_lines
            r = r + 1

        # generate vertical borders
        r = 1
        for row in self.table_cell_matrix:
            v_lines = list(map(lambda x: f"\t\t{x}", row.vertical_borders_odt(r, color_dict)))
            table_lines = table_lines + v_lines
            r = r + 1

        # close the table definition
        table_lines.append(f"\t}}")

        # generate cell values
        for row in self.table_cell_matrix:
            row_lines = list(map(lambda x: f"\t{x}", row.cell_content_odt(include_formatting=False, color_dict=color_dict, strip_comments=strip_comments, header_footer=header_footer)))
            table_lines = table_lines + row_lines

        table_lines.append(f"\t\\end{{{table_type}}}")
        table_lines = list(map(lambda x: f"\t{x}", table_lines))

        if not strip_comments:
            table_lines = [f"% OdtTable: ({self.start_row+1}-{self.end_row+1}) : {self.row_count} rows"] + table_lines

        return table_lines


''' Odt Block object wrapper
'''
class OdtParagraph(OdtBlock):

    ''' constructor
    '''
    def __init__(self, data_row, row_number):
        self.data_row = data_row
        self.row_number = row_number

    ''' generates the odt code
    '''
    def to_odt(self, longtable, color_dict, strip_comments=False):
        block_lines = []
        if not strip_comments:
            block_lines.append(f"% OdtParagraph: row {self.row_number+1}")

        # TODO 3: generate the block
        if len(self.data_row.cells) > 0:
            row_text = self.data_row.get_cell(0).content_odt(include_formatting=True, color_dict=color_dict)
            row_lines = list(map(lambda x: f"\t{x}", row_text))
            block_lines = block_lines + row_lines

        return block_lines
