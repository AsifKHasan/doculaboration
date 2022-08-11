#!/usr/bin/env python3

import json
import importlib
import inspect
from pprint import pprint
 
from latex.latex_util import *
from helper.logger import *

#   ----------------------------------------------------------------------------------------------------------------
#   latex section objects wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' Latex section base object
'''
class LatexSectionBase(object):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        self._config = config
        self._section_data = section_data

        self.section = self._section_data['section']
        self.level = self._section_data['level']
        self.page_numbering = self._section_data['hide-pageno']
        self.section_index = self._section_data['section-index']
        self.section_break = self._section_data['section-break']
        self.page_break = self._section_data['page-break']
        self.first_section = self._section_data['first-section']
        self.hide_heading = self._section_data['hide-heading']
        self.heading = section_data['heading']

        self.nesting_level = self._section_data['nesting-level']
        self.parent_section_index_text = self._section_data['parent-section-index-text']

        zfilled_index = str(self.section_index).zfill(3)
        if self.parent_section_index_text != '':
            self.section_index_text = f"{self.parent_section_index_text}_{zfilled_index}"
        else:
            self.section_index_text = zfilled_index

        self.id = f"s_{self.section_index_text}"

        self._section_data['landscape'] = 'landscape' if self._section_data['landscape'] else 'portrait'


        # master-page name
        self.landscape = self._section_data['landscape']

        self.page_spec = self._config['page-specs']['page-spec'][self._section_data['page-spec']]
        self.margin_spec = self._config['page-specs']['margin-spec'][self._section_data['margin-spec']]
        self._section_data['width'] = float(self.page_spec['width']) - float(self.margin_spec['left']) - float(self.margin_spec['right']) - float(self.margin_spec['gutter'])
        self._section_data['height'] = float(self.page_spec['height']) - float(self.margin_spec['top']) - float(self.margin_spec['bottom'])

        self.section_width = self._section_data['width']
        self.section_height = self._section_data['height']

        self.page_style_name = f"pagestyle{LETTERS[self.section_index]}"

        # headers and footers
        # print(f".. orientation: {landscape}, section_width: {self.section_width}")
        self.header_first = LatexPageHeaderFooter(content_data=section_data['header-first'], section_width=self.section_width, section_index=self.section_index, section_id=self.id, header_footer='header', odd_even='first', nesting_level=self.nesting_level)
        self.header_odd   = LatexPageHeaderFooter(content_data=section_data['header-odd'],   section_width=self.section_width, section_index=self.section_index, section_id=self.id, header_footer='header', odd_even='odd',   nesting_level=self.nesting_level)
        self.header_even  = LatexPageHeaderFooter(content_data=section_data['header-even'],  section_width=self.section_width, section_index=self.section_index, section_id=self.id, header_footer='header', odd_even='even',  nesting_level=self.nesting_level)
        self.footer_first = LatexPageHeaderFooter(content_data=section_data['footer-first'], section_width=self.section_width, section_index=self.section_index, section_id=self.id, header_footer='footer', odd_even='first', nesting_level=self.nesting_level)
        self.footer_odd   = LatexPageHeaderFooter(content_data=section_data['footer-odd'],   section_width=self.section_width, section_index=self.section_index, section_id=self.id, header_footer='footer', odd_even='odd',   nesting_level=self.nesting_level)
        self.footer_even  = LatexPageHeaderFooter(content_data=section_data['footer-even'],  section_width=self.section_width, section_index=self.section_index, section_id=self.id, header_footer='footer', odd_even='even',  nesting_level=self.nesting_level)

        self.section_contents = LatexContent(content_data=section_data.get('contents'), content_width=self.section_width, section_index=self.section_index, section_id=self.id, nesting_level=self.nesting_level)



    ''' generates section geometry
    '''
    def get_geometry(self):
        geometry_lines = []

        paper = self.page_spec['name']
        page_width = self.page_spec['width']
        page_height = self.page_spec['height']
        top_margin = self.margin_spec['top']
        bottom_margin = self.margin_spec['bottom']
        left_margin = self.margin_spec['left']
        right_margin = self.margin_spec['right']

        geometry_lines.append(f"% LatexSection: [{self.id}]")
        geometry_lines.append(f"\t\\pagebreak")

        geometry_lines.append(f"\t\\newgeometry{{{paper}, top={top_margin}in, bottom={bottom_margin}in, left={left_margin}in, right={right_margin}in, {self.landscape}}}")

        geometry_lines = mark_as_latex(lines=geometry_lines)

        return geometry_lines



    ''' generates section heading
    '''
    def get_heading(self):
        heading_lines = []
        if not self.hide_heading:
            heading_text = self.heading
            if self.section != '':
                heading_text = f"{self.section} {heading_text}".strip()

            # headings are styles based on level
            outline_level = self.level + self.nesting_level
            if outline_level == 0:
                heading_lines.append(f"\\titlestyle{{ {heading_text} }}")
            else:
                heading_tag = LATEX_HEADING_MAP.get(f"Heading {outline_level}")
                heading_lines.append(f"\\phantomsection")
                heading_lines.append(f"\\{heading_tag}{{ {heading_text} }}")

        return mark_as_latex(lines=heading_lines)



    ''' get header/fotter contents
    '''
    def get_header_footer(self, color_dict, document_footnotes):
        hf_lines = []

        # get headers and footers
        if self.header_first.has_content:
            hf_lines = hf_lines +  list(map(lambda x: f"\t\{x}", self.header_first.content_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)))

        if self.header_odd.has_content:
            hf_lines = hf_lines +  list(map(lambda x: f"\t{x}", self.header_odd.content_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)))

        if self.header_even.has_content:
            hf_lines = hf_lines +  list(map(lambda x: f"\t{x}", self.header_even.content_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)))

        if self.footer_first.has_content:
            hf_lines = hf_lines +  list(map(lambda x: f"\t{x}", self.footer_first.content_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)))

        if self.footer_odd.has_content:
            hf_lines = hf_lines +  list(map(lambda x: f"\t{x}", self.footer_odd.content_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)))

        if self.footer_even.has_content:
            hf_lines = hf_lines +  list(map(lambda x: f"\t{x}", self.footer_even.content_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)))

        # now the pagestyle
        hf_lines = hf_lines + fancy_pagestyle_header(self.page_style_name)
        if self.header_odd.has_content:
            hf_lines.append(f"\t\t\t\\fancyhead[O]{{\\{self.header_odd.id}}}")

        if self.header_even.has_content:
            hf_lines.append(f"\t\t\t\\fancyhead[E]{{\\{self.header_even.id}}}")

        if self.footer_odd.has_content:
            hf_lines.append(f"\t\t\t\\fancyfoot[O]{{\\{self.footer_odd.id}}}")

        if self.footer_even.has_content:
            hf_lines.append(f"\t\t\t\\fancyfoot[E]{{\\{self.footer_even.id}}}")

        hf_lines.append(f"\t\t}}")
        hf_lines.append(f"\t\\pagestyle{{{self.page_style_name}}}")

        # TODO
        hf_lines = [f"% PageStyle - [{self.page_style_name}]"] + hf_lines

        hf_lines = mark_as_latex(lines=hf_lines)

        return hf_lines



    ''' generates the latex code
    '''
    def section_to_latex(self, color_dict, document_footnotes):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        geometry_lines = []
        # page geometry changes only when a new section starts or if it is the first section
        if self.section_break or self.first_section:
            geometry_lines = self.get_geometry()


        # header/footer lines
        hf_lines = self.get_header_footer(color_dict=color_dict, document_footnotes=document_footnotes)


        content_lines = []
        # the section may have a page-break
        if self.page_break:
            content_lines.append(f"\t\\pagebreak")

        content_lines = mark_as_latex(lines=content_lines)


        # section heading is always applicable
        heading_lines = self.get_heading()

        return geometry_lines + hf_lines + content_lines + heading_lines



''' Latex table section object
'''
class LatexTableSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        super().__init__(section_data=section_data, config=config)

 
    ''' generates the latex code
    '''
    def section_to_latex(self, color_dict, document_footnotes):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        section_lines = super().section_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)

        # get the contents
        content_lines = self.section_contents.content_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)

        return section_lines + content_lines



''' Latex gsheet section object
'''
class LatexGsheetSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        super().__init__(section_data=section_data, config=config)


    ''' generates the odt code
    '''
    def section_to_latex(self, color_dict, document_footnotes):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        section_lines = super().section_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)

        # for embedded gsheets, 'contents' does not contain the actual content to render, rather we get a list of sections where each section contains the actual content
        if self._section_data['contents'] is not None and 'sections' in self._section_data['contents']:
            # these are child contents, we need to assign indexes so that they do not overlap with parent indexes
            nesting_level = self.nesting_level + 1

            first_section = False
            section_index = self.section_index * 100
            for section in self._section_data['contents']['sections']:
                section['nesting-level'] = nesting_level
                section['parent-section-index-text'] = self.section_index_text
                if section['section'] != '':
                    info(msg=f"writing : {section['section'].strip()} {section['heading'].strip()}", nesting_level=nesting_level)
                else:
                    info(msg=f"writing : {section['heading'].strip()}", nesting_level=nesting_level)

                section['first-section'] = True if first_section else False
                section['section-index'] = section_index

                module = importlib.import_module("latex.latex_api")
                func = getattr(module, f"process_{section['content-type']}")
                section_lines = section_lines + func(section_data=section, config=self._config, color_dict=color_dict)

                first_section = False
                section_index = section_index + 1

        return section_lines



''' Latex ToC section object
'''
class LatexToCSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data=section_data, config=config)


    ''' generates the latex code
    '''
    def section_to_latex(self, color_dict, document_footnotes):
        section_lines = super().section_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)

        # table-of-contents
        content_lines = []
        content_lines.append("\\renewcommand{\\contentsname}{}")
        content_lines.append("\\vspace{-0.5in}")
        content_lines.append("\\tableofcontents")
        content_lines.append("\\addtocontents{toc}{~\\hfill\\textbf{Page}\\par}")
        content_lines = mark_as_latex(lines=content_lines)

        return section_lines + content_lines



''' Latex LoT section object
'''
class LatexLoTSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data=section_data, config=config)


    ''' generates the latex code
    '''
    def section_to_latex(self, color_dict, document_footnotes):
        section_lines = super().section_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)

        # table-of-contents
        content_lines = []
        content_lines.append("\\renewcommand{\\listtablename}{}")
        content_lines.append("\\vspace{-0.4in}")
        content_lines.append("\\listoftables")
        content_lines.append("\\addtocontents{lot}{~\\hfill\\textbf{Page}\\par}")
        content_lines = mark_as_latex(lines=content_lines)

        return section_lines + content_lines



''' Latex LoF section object
'''
class LatexLoFSection(LatexSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data=section_data, config=config)


    ''' generates the latex code
    '''
    def section_to_latex(self, color_dict, document_footnotes):
        section_lines = super().section_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)

        # table-of-contents
        content_lines = []
        content_lines.append("\\renewcommand{\\listfigurename}{}")
        content_lines.append("\\vspace{-0.4in}")
        content_lines.append("\\listoffigures")
        content_lines.append("\\addtocontents{lof}{~\\hfill\\textbf{Page}\\par}")
        content_lines = mark_as_latex(lines=content_lines)

        return section_lines + content_lines



''' Latex section content base object
'''
class LatexContent(object):

    ''' constructor
    '''
    def __init__(self, content_data, content_width, section_index, section_id, nesting_level):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        self.content_width, self.section_index, self.section_id, self.nesting_level = content_width, section_index, section_id, nesting_level

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
            # self.default_format = CellFormat(format_dict=properties.get('defaultFormat'))

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
                        self.row_metadata_list.append(RowMetadata(row_metadata_dict=row_metadata))

                    # columnMetadata
                    for column_metadata in data.get('columnMetadata', []):
                        self.column_metadata_list.append(ColumnMetadata(column_metadata_dict=column_metadata))

                    # merges
                    for merge in sheets[0].get('merges', []):
                        self.merge_list.append(Merge(gsheet_merge_dict=merge, start_row=self.start_row, start_column=self.start_column))

                    # column width needs adjustment as \tabcolsep is COLSEPin. This means each column has a COLSEP inch on left and right as space which needs to be removed from column width
                    all_column_widths_in_pixel = sum(x.pixel_size for x in self.column_metadata_list)
                    self.column_widths = [ (x.pixel_size * self.content_width / all_column_widths_in_pixel) - (COLSEP * 2) for x in self.column_metadata_list ]

                    # rowData
                    r = 0
                    for row_data in data.get('rowData', []):
                        row = Row(row_num=r, row_data=row_data, default_format=self.default_format, section_width=self.content_width, column_widths=self.column_widths, row_height=self.row_metadata_list[r].inches, nesting_level=self.nesting_level)
                        self.cell_matrix.append(row)
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
            first_cell = first_row_object.get_cell(c=first_col)

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
                    next_cell_in_row = next_row_object.get_cell(c=c)

                    if next_cell_in_row is None:
                        # the cell may not be existing at all, we have to create
                        # debug(f"..cell [{r+1},{c+1}] does not exist, to be inserted")
                        next_cell_in_row = Cell(row_num=r, col_num=c, value=None, default_format=first_cell.default_format, column_widths=first_cell.column_widths, row_height=row_height)
                        next_row_object.insert_cell(pos=c, cell=next_cell_in_row)

                    if next_cell_in_row.is_empty:
                        # debug(f"..cell [{r+1},{c+1}] is empty")
                        # the cell is a newly inserted one, its format should be the same (for borders, colors) as the first cell so that we can draw borders properly
                        next_cell_in_row.copy_format_from(from_cell=first_cell)

                        # mark cells for multicol only if it is multicol
                        if col_span > 1:
                            if c == first_col:
                                # the last cell of the merge to be marked as LastCell
                                # debug(f"..cell [{r+1},{c+1}] is the LastCell of the column merge")
                                next_cell_in_row.mark_multicol(span=MultiSpan.FirstCell)

                            elif c == last_col-1:
                                # the last cell of the merge to be marked as LastCell
                                # debug(f"..cell [{r+1},{c+1}] is the LastCell of the column merge")
                                next_cell_in_row.mark_multicol(span=MultiSpan.LastCell)

                            else:
                                # the inner cells of the merge to be marked as InnerCell
                                # debug(f"..cell [{r+1},{c+1}] is an InnerCell of the column merge")
                                next_cell_in_row.mark_multicol(span=MultiSpan.InnerCell)

                        else:
                            next_cell_in_row.mark_multicol(span=MultiSpan.No)


                        # mark cells for multirow only if it is multirow
                        if row_span > 1:
                            if r == first_row:
                                # the last cell of the merge to be marked as LastCell
                                # debug(f"..cell [{r+1},{c+1}] is the LastCell of the row merge")
                                next_cell_in_row.mark_multirow(span=MultiSpan.FirstCell)

                            elif r == last_row-1:
                                # the last cell of the merge to be marked as LastCell
                                # debug(f"..cell [{r+1},{c+1}] is the LastCell of the row merge")
                                next_cell_in_row.mark_multirow(span=MultiSpan.LastCell)

                            else:
                                # the inner cells of the merge to be marked as InnerCell
                                # debug(f"..cell [{r+1},{c+1}] is an InnerCell of the row merge")
                                next_cell_in_row.mark_multirow(span=MultiSpan.InnerCell)

                        else:
                            next_cell_in_row.mark_multirow(span=MultiSpan.No)

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
            if data_row.is_free_content():
                # there may be a pending/running table
                if r > next_table_starts_in_row:
                    table = LatexTable(cell_matrix=self.cell_matrix, start_row=next_table_starts_in_row, end_row=(r - 1), column_widths=self.column_widths, section_index=self.section_index, section_id=self.section_id)
                    self.content_list.append(table)

                block = LatexParagraph(data_row=data_row, row_number=r, section_index=self.section_index, section_id=self.section_id)
                self.content_list.append(block)

                next_table_starts_in_row = r + 1

            # the row may start with a note of repeat-rows which means that a new table is atarting
            elif data_row.is_table_start():
                # there may be a pending/running table
                if r > next_table_starts_in_row:
                    table = LatexTable(cell_matrix=self.cell_matrix, start_row=next_table_starts_in_row, end_row=(r - 1), column_widths=self.column_widths, section_index=self.section_index, section_id=self.section_id)
                    self.content_list.append(table)

                    next_table_starts_in_row = r

            else:
                next_table_ends_in_row = r

        # there may be a pending/running table
        if next_table_ends_in_row >= next_table_starts_in_row:
            table = LatexTable(cell_matrix=self.cell_matrix, start_row=next_table_starts_in_row, end_row=next_table_ends_in_row, column_widths=self.column_widths, section_index=self.section_index, section_id=self.section_id)
            self.content_list.append(table)


    ''' generates the latex code
    '''
    def content_to_latex(self, color_dict, document_footnotes):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        latex_lines = []

        # iterate through tables and blocks contents
        last_block_is_a_paragraph = False
        for block in self.content_list:
            # if current block is a paragraph and last block was also a paragraph, insert a newline
            if isinstance(block, LatexParagraph) and last_block_is_a_paragraph:
                latex_lines.append(f"\t\\\\[{0}pt]")

            latex_lines = latex_lines + block.block_to_latex(longtable=True, color_dict=color_dict, document_footnotes=document_footnotes)

            # keep track of the block as the previous block
            if isinstance(block, LatexParagraph):
                last_block_is_a_paragraph = True
            else:
                last_block_is_a_paragraph = False

        latex_lines = mark_as_latex(lines=latex_lines)

        return latex_lines



''' Page Header/Footer
'''
class LatexPageHeaderFooter(LatexContent):

    ''' constructor
        header_footer : header/footer
        odd_even      : first/odd/even
    '''
    def __init__(self, content_data, section_width, section_index, section_id, header_footer, odd_even, nesting_level):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        super().__init__(content_data=content_data, content_width=section_width, section_index=section_index, section_id=section_id, nesting_level=nesting_level)
        self.header_footer, self.odd_even = header_footer, odd_even
        self.id = f"{self.header_footer}{self.odd_even}{LETTERS[section_index]}"


    ''' generates the latex code
    '''
    def content_to_latex(self, color_dict, document_footnotes):
        latex_lines = []

        latex_lines.append(f"\\newcommand{{\\{self.id}}}{{%")

        # iterate through tables and blocks contents
        first_block = True
        for block in self.content_list:
            block_lines = block.block_to_latex(longtable=False, color_dict=color_dict, document_footnotes=document_footnotes, strip_comments=True, header_footer=self.header_footer)
            # block_lines = list(map(lambda x: f"\t{x}", block_lines))
            latex_lines = latex_lines + block_lines

            first_block = False

        latex_lines.append(f"\t}}")

        return [f"% LatexPageHeaderFooter: [{self.id}]"] + latex_lines



''' Latex Block object wrapper base class (plain latex, table, header etc.)
'''
class LatexBlock(object):

    ''' constructor
    '''
    def __init__(self, section_index, section_id):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        self.section_index, self.section_id = section_index, section_id



    ''' generates latex code
    '''
    def block_to_latex(self, longtable, color_dict, document_footnotes, strip_comments=False, header_footer=None):
        pass



''' Latex Table object wrapper
'''
class LatexTable(LatexBlock):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, cell_matrix, start_row, end_row, column_widths):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        super().__init__(section_index, section_id)

        self.start_row, self.end_row, self.column_widths = start_row, end_row, column_widths
        self.table_cell_matrix = cell_matrix[start_row:end_row+1]
        self.row_count = len(self.table_cell_matrix)
        self.table_name = f"LatexTable: {self.start_row+1}-{self.end_row+1}[{self.row_count}]"
        self.id = f"{self.section_id}__t_{self.start_row+1}_{self.end_row+1}"

        # header row if any
        self.header_row_count = self.table_cell_matrix[0].get_cell(c=0).note.header_rows


    ''' generates the latex code
    '''
    def block_to_latex(self, longtable, color_dict, document_footnotes, strip_comments=False, header_footer=None):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        # for storing the footnotes (if any) for this block
        if (self.id not in document_footnotes):
            document_footnotes[self.id] = []

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
            row_format_line = f"\t\t{row.row_format_to_latex(r=r, color_dict=color_dict)}"
            table_lines.append(row_format_line)
            r = r + 1

        # generate cell formats
        r = 1
        for row in self.table_cell_matrix:
            cell_format_lines = list(map(lambda x: f"\t\t{x}", row.cell_formats_to_latex(r=r, color_dict=color_dict)))
            table_lines = table_lines + cell_format_lines
            r = r + 1

        # generate vertical borders
        r = 1
        for row in self.table_cell_matrix:
            v_lines = list(map(lambda x: f"\t\t{x}", row.vertical_borders_to_latex(r=r, color_dict=color_dict)))
            table_lines = table_lines + v_lines
            r = r + 1

        # close the table definition
        table_lines.append(f"\t}}")

        # generate cell values
        for row in self.table_cell_matrix:
            row_lines = list(map(lambda x: f"\t{x}", row.cell_contents_to_latex(block_id=self.id, include_formatting=False, color_dict=color_dict, document_footnotes=document_footnotes, strip_comments=strip_comments, header_footer=header_footer)))
            table_lines = table_lines + row_lines

        table_lines.append(f"\t\\end{{{table_type}}}")
        table_lines = list(map(lambda x: f"\t{x}", table_lines))

        # latex footnotes
        footnote_texts = document_footnotes[self.id]
        if len(footnote_texts):
            # append footnotetexts 
            table_lines.append(f"")
            for footnote_text_dict in footnote_texts:
                footnote_text = f"\t\\footnotetext[{footnote_text_dict['mark']}]{{{footnote_text_dict['text']}}}"
                table_lines.append(footnote_text)

            # \\setfnsymbol for the footnotes, this needs to go before the table
            table_lines = [f"\\setfnsymbol{{{self.id}_symbols}}"] + table_lines

        if not strip_comments:
            table_lines = [f"% LatexTable: ({self.start_row+1}-{self.end_row+1}) : {self.row_count} rows"] + table_lines

        return [f"% LatexTable: [{self.id}]"] + table_lines



''' Latex Block object wrapper
'''
class LatexParagraph(LatexBlock):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, data_row, row_number):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        super().__init__(section_index, section_id)

        self.data_row = data_row
        self.row_number = row_number
        self.id = f"{self.section_id}__p_{self.row_number+1}"


    ''' generates the latex code
    '''
    def block_to_latex(self, longtable, color_dict, document_footnotes, strip_comments=False):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        # for storing the footnotes (if any) for this block
        if (self.id not in document_footnotes):
            document_footnotes[self.id] = []

        block_lines = []
        if not strip_comments:
            block_lines.append(f"% LatexParagraph: row {self.row_number+1}")

        # generate the block, only the first cell of the data_row to be produced
        if len(self.data_row.cells) > 0:
            row_text = self.data_row.get_cell(c=0).cell_content_to_latex(block_id=self.id, include_formatting=True, color_dict=color_dict, document_footnotes=document_footnotes)
            row_lines = list(map(lambda x: f"\t{x}", row_text))
            block_lines = block_lines + row_lines

        # latex footnotes
        footnote_texts = document_footnotes[self.id]
        if len(footnote_texts):
            # append footnotetexts 
            block_lines.append(f"")
            for footnote_text_dict in footnote_texts:
                footnote_text = f"\t\\footnotetext[{footnote_text_dict['mark']}]{{{footnote_text_dict['text']}}}"
                block_lines.append(footnote_text)

            block_lines.append(f"")

            # \\setfnsymbol for the footnotes, this needs to go before the table
            block_lines = [f"\\setfnsymbol{{{self.id}_symbols}}"] + block_lines

        return [f"% LatexParagraph: [{self.id}]"] + block_lines



#   ----------------------------------------------------------------------------------------------------------------
#   gsheet cell wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' gsheet Cell object wrapper
'''
class Cell(object):

    ''' constructor
    '''
    def __init__(self, row_num, col_num, value, default_format, column_widths, row_height, nesting_level):
        self.row_num, self.col_num, self.column_widths, self.default_format, self.nesting_level = row_num, col_num, column_widths, default_format, nesting_level
        self.cell_name = f"cell: [{self.row_num},{self.col_num}]"
        self.value = value
        self.text_format_runs = []
        self.cell_width = self.column_widths[self.col_num]
        self.cell_height = row_height
        self.merge_spec = CellMergeSpec()

        if self.value:
            self.note = CellNote(note_json=value.get('note'))
            self.formatted_value = self.value.get('formattedValue', '')

            # self.effective_format = CellFormat(format_dict=self.value.get('effectiveFormat'), default_format=self.default_format)
            self.effective_format = CellFormat(format_dict=self.value.get('effectiveFormat'))

            for text_format_run in self.value.get('textFormatRuns', []):
                self.text_format_runs.append(TextFormatRun(run_dict=text_format_run, default_format=self.effective_format.text_format.source))

            # presence of userEnteredFormat makes the cell non-empty
            if 'userEnteredFormat' in self.value:
                self.user_entered_format = CellFormat(format_dict=self.value.get('userEnteredFormat'))
                self.is_empty = False
            else:
                self.user_entered_format = None
                self.is_empty = True


            # we need to identify exactly what kind of value the cell contains
            if 'contents' in self.value:
                self.cell_value = ContentValue(effective_format=self.effective_format, content_value=self.value['contents'])

            elif 'userEnteredValue' in self.value:
                if 'image' in self.value['userEnteredValue']:
                    self.cell_value = ImageValue(effective_format=self.effective_format, image_value=self.value['userEnteredValue']['image'])

                else:
                    if len(self.text_format_runs):
                        self.cell_value = TextRunValue(effective_format=self.effective_format, text_format_runs=self.text_format_runs, formatted_value=self.formatted_value)

                    elif self.note.page_number:
                        self.cell_value = PageNumberValue(effective_format=self.effective_format, short=False)

                    else:
                        self.cell_value = StringValue(effective_format=self.effective_format, string_value=self.value['userEnteredValue'], formatted_value=self.formatted_value, nesting_level=self.nesting_level, outline_level=self.note.outline_level)

            else:
                # self.cell_value = StringValue(effective_format=self.effective_format, '', self.formatted_value)
                # self.cell_value = None
                # warn(f"{self} is None")
                self.cell_value = StringValue(effective_format=self.effective_format, string_value='', formatted_value=self.formatted_value, nesting_level=self.nesting_level, outline_level=self.note.outline_level)



        else:
            # value can have a special case it can be an empty ditionary when the cell is an inner cell of a column merge
            self.merge_spec.multi_col = MultiSpan.No
            self.note = CellNote()
            self.cell_value = None
            self.formatted_value = None
            self.effective_format = None
            self.user_entered_format = None
            self.is_empty = True
            # warn(f"{self} is Empty")


    ''' string representation
    '''
    def __repr__(self):
        # s = f"[{self.row_num+1},{self.col_num+1:>2}], value: {not self.is_empty:<1}, mr: {self.merge_spec.multi_row:<9}, mc: {self.merge_spec.multi_col:<9} [{self.formatted_value}]"
        s = f"[{self.row_num+1},{self.col_num+1:>2}], width: {self.cell_width}]"
        return s


    ''' latex code for cell content
    '''
    def cell_content_to_latex(self, block_id, include_formatting, color_dict, document_footnotes, strip_comments=False, left_hspace=None, right_hspace=None):
        content_lines = []

        if not strip_comments:
            content_lines.append(f"% {self.merge_spec.to_string()}")

        # the content is not valid for multirow LastCell and InnerCell
        if self.merge_spec.multi_row in [MultiSpan.No, MultiSpan.FirstCell] and self.merge_spec.multi_col in [MultiSpan.No, MultiSpan.FirstCell]:
            if self.cell_value:

                # get the latex code
                cell_value = self.cell_value.value_to_latex(block_id=block_id, container_width=self.cell_width, container_height=self.cell_height, color_dict=color_dict, document_footnotes=document_footnotes, footnote_list=self.note.footnotes)

                # paragraphs need formatting to be included, table cells do not need them
                if include_formatting:
                    # alignments and bgcolor
                    if self.effective_format:
                        halign = PARA_HALIGN.get(self.effective_format.halign.halign)
                    else:
                        halign = PARA_HALIGN.get('LEFT')

                    cell_value = f"{halign}{{{cell_value}}}"


                # the cell may have a left_hspace or right_hspace
                if left_hspace:
                    cell_value = f"\\hspace{{{left_hspace}pt}}{cell_value}"

                if right_hspace:
                    cell_value = f"{cell_value}\\hspace{{{right_hspace}pt}}"

                # handle new-page defined in notes
                if self.note.new_page:
                    content_lines.append(f"\\pagebreak")

                # handle keep-with-previous defined in notes
                if self.note.keep_with_previous:
                    content_lines.append(f"\\nopagebreak")

                # the actual content
                content_lines.append(cell_value)

                # handle styles defined in notes
                if self.note.style == 'Figure':
                    # caption for figure
                    # content_lines.append(f"\\neeedspace{{2em}}")
                    content_lines.append(f"\\phantomsection")
                    content_lines.append(f"\\addcontentsline{{lof}}{{figure}}{{{tex_escape(self.user_entered_value.string_value)}}}")
                elif self.note.style == 'Table':
                    # caption for table
                    # content_lines.append(f"\\neeedspace{{2em}}")
                    content_lines.append(f"\\phantomsection")
                    content_lines.append(f"\\addcontentsline{{lot}}{{table}}{{{tex_escape(self.user_entered_value.string_value)}}}")
                elif self.note.style:
                    # some custom style needs to be applied
                    heading_tag = LATEX_HEADING_MAP.get(self.note.style)
                    if heading_tag:
                        # content_lines.append(f"\\neeedspace{{2em}}")
                        content_lines.append(f"\\phantomsection")
                        content_lines.append(f"\\addcontentsline{{toc}}{{{heading_tag}}}{{{tex_escape(self.user_entered_value.string_value)}}}")
                    else:
                        warn(f"style : {self.note.style} not defined")

        return content_lines


    ''' latex code for cell format
    '''
    def cell_format_to_latex(self, r, color_dict):
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


    ''' Copy format from the cell passed
    '''
    def copy_format_from(self, from_cell):
        self.user_entered_format = from_cell.user_entered_format
        self.effective_format = from_cell.effective_format

        self.merge_spec.multi_col = from_cell.merge_spec.multi_col
        self.merge_spec.col_span = from_cell.merge_spec.col_span
        self.merge_spec.row_span = from_cell.merge_spec.row_span
        self.cell_width = from_cell.cell_width


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
    def top_border_to_latex(self, color_dict):
        t = None
        if self.effective_format and self.effective_format.borders:
            if self.top_border_allowed():
                t = self.effective_format.borders.borders_to_latex_t(color_dict=color_dict)
                if t is not None:
                    # t = f"{{{self.col_num+1}-{self.col_num+self.merge_spec.col_span}}}{{{t}}}"
                    t = f"[{t}]{{{self.col_num+1}-{self.col_num+self.merge_spec.col_span}}}"

        return t


    ''' latex code for bottom border
    '''
    def bottom_border_to_latex(self, color_dict):
        b = None
        if self.effective_format and self.effective_format.borders:
            if self.bottom_border_allowed():
                b = self.effective_format.borders.borders_to_latex_b(color_dict=color_dict)
                if b is not None:
                    # b = f"{{{self.col_num+1}-{self.col_num+self.merge_spec.col_span}}}{{{b}}}"
                    b = f"[{b}]{{{self.col_num+1}-{self.col_num+self.merge_spec.col_span}}}"

        return b


    ''' latex code for left and right borders
        r is row numbner (1 based)
    '''
    def cell_vertical_borders_to_latex(self, r, color_dict):
        lr_borders = []
        if self.effective_format and self.effective_format.borders:
            if self.left_border_allowed():
                lb = self.effective_format.borders.borders_to_latex_l(color_dict=color_dict)
                if lb is not None:
                    lb = f"vline{{{self.col_num+1}}} = {{{r}}}{{{lb}}},"
                    lr_borders.append(lb)

            if self.right_border_allowed():
                rb = self.effective_format.borders.borders_to_latex_r(color_dict=color_dict)
                if rb is not None:
                    rb = f"vline{{{self.col_num+2}}} = {{{r}}}{{{rb}}},"
                    lr_borders.append(rb)

        return lr_borders



''' gsheet Row object wrapper
'''
class Row(object):

    ''' constructor
    '''
    def __init__(self, row_num, row_data, default_format, section_width, column_widths, row_height, nesting_level):
        self.row_num, self.default_format, self.section_width, self.column_widths, self.row_height, self.nesting_level = row_num, default_format, section_width, column_widths, row_height, nesting_level
        self.row_name = f"row: [{self.row_num+1}]"

        self.cells = []
        c = 0
        for value in row_data.get('values', []):
            cell = Cell(row_num=self.row_num, col_num=c, value=value, default_format=self.default_format, column_widths=self.column_widths, row_height=self.row_height, nesting_level=self.nesting_level)
            self.cells.append(cell)
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


    ''' it is true only when the first cell has a free_content true value
    '''
    def is_free_content(self):
        if len(self.cells) > 0:
            # the first cell is the relevant cell only
            if self.cells[0]:
                return self.cells[0].note.free_content
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
    def top_borders_to_latex(self, color_dict):
        top_borders = []
        c = 0
        for cell in self.cells:
            if cell is None:
                warn(f"{self.row_name} has a Null cell at {c}")

            else:
                t = cell.top_border_to_latex(color_dict=color_dict)
                if t is not None:
                    # top_borders.append(f"\\SetHline{t}")
                    top_borders.append(f"\\cline{t}")

            c = c + 1

        return top_borders


    ''' generates the bottom borders
    '''
    def bottom_borders_to_latex(self, color_dict):
        bottom_borders = []
        c = 0
        for cell in self.cells:
            if cell is None:
                warn(f"{self.row_name} has a Null cell at {c}")

            else:
                b = cell.bottom_border_to_latex(color_dict=color_dict)
                if b is not None:
                    # bottom_borders.append(f"\\SetHline{b}")
                    bottom_borders.append(f"\\cline{b}")

            c = c + 1

        return bottom_borders


    ''' generates the vertical borders
    '''
    def vertical_borders_to_latex(self, r, color_dict):
        v_lines = []
        c = 0
        for cell in self.cells:
            if cell is not None:
                v_lines = v_lines + cell.cell_vertical_borders_to_latex(r=r, color_dict=color_dict)

            c = c + 1

        return v_lines


    ''' generates the latex code for row formats
    '''
    def row_format_to_latex(self, r, color_dict):
        row_format_line = f"row{{{r}}} = {{ht={self.row_height}in}},"

        return row_format_line


    ''' generates the latex code for cell formats
    '''
    def cell_formats_to_latex(self, r, color_dict):
        cell_format_lines = []
        for cell in self.cells:
            if cell is not None:
                if not cell.is_empty:
                    cell_format_lines = cell_format_lines + cell.cell_format_to_latex(r=r, color_dict=color_dict)

        return cell_format_lines


    ''' generates the latex code
    '''
    def cell_contents_to_latex(self, block_id, include_formatting, color_dict, document_footnotes, strip_comments=False, header_footer=None):
        # debug(f"processing {self.row_name}")

        row_lines = []
        # if not strip_comments:
        row_lines.append(f"% {self.row_name}")

        # borders
        top_border_lines = self.top_borders_to_latex(color_dict=color_dict)
        top_border_lines = list(map(lambda x: f"\t{x}", top_border_lines))

        bottom_border_lines = self.bottom_borders_to_latex(color_dict=color_dict)
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
                left_hspace = None
                right_hspace = None
                # if the cell is for a header/footer based on column add hspace
                if header_footer in ['header', 'footer']:
                    # first column has a left -ve hspace
                    if c == 0:
                        left_hspace = HEADER_FOOTER_FIRST_COL_HSPACE

                    # first column has a left -ve hspace
                    if c == len(self.cells) - 1:
                        right_hspace = HEADER_FOOTER_LAST_COL_HSPACE

                cell_lines = cell.cell_content_to_latex(block_id=block_id, include_formatting=include_formatting, color_dict=color_dict, document_footnotes=document_footnotes, strip_comments=strip_comments, left_hspace=left_hspace, right_hspace=right_hspace)

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
            self.fgcolor = RgbColor(rgb_dict=text_format_dict.get('foregroundColor'))
            if 'fontFamily' in text_format_dict:
                font_family = text_format_dict['fontFamily']

                if font_family == DEFAULT_FONT:
                    # it is the default font, so we do not need to set it
                    self.font_family = ''

                else:
                    if font_family in FONT_MAP:
                        # we have the font in our FONT_MAP, we replace it as per the map
                        self.font_family = FONT_MAP.get(font_family)
                        if self.font_family != font_family:
                            warn(f"{font_family} is mapped to {self.font_family}")

                    else:
                        # the font is not in our FONT_MAP, use it anyway and let us see
                        self.font_family = font_family

                        # self.font_family = DEFAULT_FONT
                        # warn(f"{text_format_dict['fontFamily']} is not mapped, will use default font")
            else:
                self.font_family = ''

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


    ''' generate latex code
    '''
    def text_format_to_latex(self, block_id, text, color_dict, document_footnotes, footnote_list, verbatim=False):
        color_dict[self.fgcolor.key()] = self.fgcolor.value()

        # process footnote (if any)
        content = f"{process_footnotes(block_id=block_id, text_content=text, document_footnotes=document_footnotes, footnote_list=footnote_list, verbatim=verbatim)}"

        styled = False
        if self.is_underline:
            content = f"\\underline{{{content}}}"
            styled = True

        if self.is_strikethrough:
            content = f"\\sout{{{content}}}"
            styled = True

        if self.is_italic:
            content = f"\\textit{{{content}}}"
            styled = True

        if self.is_bold:
            content = f"\\textbf{{{content}}}"
            styled = True

        if not styled:
            content = f"{{{content}}}"

        # color, font, font-size
        if self.font_family != '':
            font_spec = f"\\fontsize{{{self.font_size}pt}}{{{self.font_size + 2}pt}}\\fontspec{{{self.font_family}}}\\color{{{self.fgcolor.key()}}}"
        else:
            font_spec = f"\\fontsize{{{self.font_size}pt}}{{{self.font_size + 2}pt}}\\color{{{self.fgcolor.key()}}}"
            # font_spec = f"\\fontsize{{{self.font_size}pt}}{{{self.font_size}pt}}\\color{{{self.fgcolor.key()}}}"

        latex = f"{font_spec}{content}"
        return latex



''' gsheet cell value object wrapper
'''
class CellValue(object):

    ''' constructor
    '''
    def __init__(self, effective_format, nesting_level=0, outline_level=0):
        self.effective_format = effective_format
        self.nesting_level = nesting_level
        self.outline_level = outline_level



''' string type CellValue
'''
class StringValue(CellValue):

    ''' constructor
    '''
    def __init__(self, effective_format, string_value, formatted_value, nesting_level=0, outline_level=0):
        super().__init__(effective_format=effective_format, nesting_level=nesting_level, outline_level=outline_level)
        if formatted_value:
            self.value = formatted_value
        else:
            if string_value and 'stringValue' in string_value:
                self.value = string_value['stringValue']
            else:
                self.value = ''


    ''' string representation
    '''
    def __repr__(self):
        s = f"string : [{self.value}]"
        return s


    ''' generates the latex code
    '''
    def value_to_latex(self, block_id, container_width, container_height, color_dict, document_footnotes, footnote_list):
        verbatim = False
        latex = self.effective_format.text_format.text_format_to_latex(block_id=block_id, text=self.value, color_dict=color_dict, document_footnotes=document_footnotes, footnote_list=footnote_list, verbatim=verbatim)

        return latex



''' text-run type CellValue
'''
class TextRunValue(CellValue):

    ''' constructor
    '''
    def __init__(self, effective_format, text_format_runs, formatted_value, nesting_level=0, outline_level=0):
        super().__init__(effective_format=effective_format, nesting_level=nesting_level, outline_level=outline_level)
        self.text_format_runs = text_format_runs
        self.formatted_value = formatted_value


    ''' string representation
    '''
    def __repr__(self):
        s = f"text-run : [{self.formatted_value}]"
        return s


    ''' generates the latex code
    '''
    def value_to_latex(self, block_id, container_width, container_height, color_dict, document_footnotes, footnote_list):
        verbatim = False
        run_value_list = []
        processed_idx = len(self.formatted_value)
        for text_format_run in reversed(self.text_format_runs):
            text = self.formatted_value[:processed_idx]
            run_value_list.insert(0, text_format_run.text_format_run_to_latex(block_id=block_id, text=text, color_dict=color_dict, document_footnotes=document_footnotes, footnote_list=footnote_list, verbatim=verbatim))
            processed_idx = text_format_run.start_index

        return ''.join(run_value_list)



''' page-number type CellValue
'''
class PageNumberValue(CellValue):

    ''' constructor
    '''
    def __init__(self, effective_format, short=False, nesting_level=0, outline_level=0):
        super().__init__(effective_format=effective_format, nesting_level=nesting_level, outline_level=outline_level)
        self.short = short


    ''' string representation
    '''
    def __repr__(self):
        s = f"page-number"
        return s


    ''' generates the latex code
    '''
    def value_to_latex(self, block_id, container_width, container_height, color_dict, document_footnotes, footnote_list):
        if self.short:
            latex = "\\thepage"
        else:
            latex = "\\thepage\\ of \\pageref{LastPage}"

        return latex



''' image type CellValue
'''
class ImageValue(CellValue):

    ''' constructor
    '''
    def __init__(self, effective_format, image_value, nesting_level=0, outline_level=0):
        super().__init__(effective_format=effective_format, nesting_level=nesting_level, outline_level=outline_level)
        self.value = image_value


    ''' string representation
    '''
    def __repr__(self):
        s = f"image : [{self.value}]"
        return s


    ''' generates the latex code
    '''
    def value_to_latex(self, block_id, container_width, container_height, color_dict, document_footnotes, footnote_list):
        # even now the width may exceed actual cell width, we need to adjust for that
        dpi_x = 96 if self.value['dpi'][0] == 0 else self.value['dpi'][0]
        dpi_y = 96 if self.value['dpi'][1] == 0 else self.value['dpi'][1]
        image_width_in_pixel = self.value['size'][0]
        image_height_in_pixel = self.value['size'][1]
        image_width_in_inches =  image_width_in_pixel / dpi_x
        image_height_in_inches = image_height_in_pixel / dpi_y

        if self.value['mode'] == 1:
            # image is to be scaled within the cell width and height
            if image_width_in_inches > container_height:
                adjust_ratio = (container_height / image_width_in_inches)
                image_width_in_inches = image_width_in_inches * adjust_ratio
                image_height_in_inches = image_height_in_inches * adjust_ratio

            if image_height_in_inches > container_height:
                # debug(f"image : [{image_width_in_inches}in X {image_height_in_inches}in, cell-width [{container_height}in], cell-height [{container_height}in]")
                adjust_ratio = (container_height / image_height_in_inches)
                image_width_in_inches = image_width_in_inches * adjust_ratio
                image_height_in_inches = image_height_in_inches * adjust_ratio

        elif self.value['mode'] == 3:
            # image size is unchanged
            pass

        else:
            # treat it as if image mode is 3
            pass

        latex = f"\includegraphics[width={image_width_in_inches}in]{{{os_specific_path(self.value['path'])}}}"

        return latex



''' content type CellValue
'''
class ContentValue(CellValue):

    ''' constructor
    '''
    def __init__(self, effective_format, content_value, nesting_level=0, outline_level=0):
        super().__init__(effective_format=effective_format, nesting_level=nesting_level, outline_level=outline_level)
        self.value = content_value


    ''' string representation
    '''
    def __repr__(self):
        s = f"content : [{self.value['sheets'][0]['properties']['title']}]"
        return s


    ''' generates the latex code
    '''
    def value_to_latex(self, block_id, container_width, container_height, color_dict, document_footnotes, footnote_list):
        section_contents = LatexContent(content_data=self.value, content_width=container_width, section_index=self.section_index, nesting_level=self.nesting_level)
        return section_contents.content_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)



''' gsheet cell format object wrapper
'''
class CellFormat(object):

    ''' constructor
    '''
    def __init__(self, format_dict, default_format=None):
        if format_dict:
            self.bgcolor = RgbColor(rgb_dict=format_dict.get('backgroundColor'))
            self.borders = Borders(borders_dict=format_dict.get('borders'))
            self.padding = Padding(padding_dict=format_dict.get('padding'))
            self.halign = HorizontalAlignment(halign=format_dict.get('horizontalAlignment'))
            self.valign = VerticalAlignment(valign=format_dict.get('verticalAlignment'))
            self.text_format = TextFormat(text_format_dict=format_dict.get('textFormat'))
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
            self.borders = Borders(borders_dict=None)
        else:
            self.borders.override_top_border(border_color=color)
            self.borders.override_bottom_border(border_color=color)
            self.borders.override_left_border(border_color=color)
            self.borders.override_right_border(border_color=color)


    ''' recolor top border with the specified color
    '''
    def recolor_top_border(self, color):
        self.borders.override_top_border(border_color=color, forced=True)


    ''' recolor bottom border with the specified color
    '''
    def recolor_bottom_border(self, color):
        self.borders.override_bottom_border(border_color=color, forced=True)



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
                self.top = Border(border_dict=borders_dict.get('top'))

            if 'right' in borders_dict:
                self.right = Border(border_dict=borders_dict.get('right'))

            if 'bottom' in borders_dict:
                self.bottom = Border(border_dict=borders_dict.get('bottom'))

            if 'left' in borders_dict:
                self.left = Border(border_dict=borders_dict.get('left'))


    ''' string representation
    '''
    def __repr__(self):
        return f"t: [{self.top}], b: [{self.bottom}], l: [{self.left}], r: [{self.right}]"


    ''' override top border with the specified color
    '''
    def override_top_border(self, border_color, forced=False):
        if border_color:
            if self.top is None:
                self.top = Border(border_dict=None)
            elif forced:
                self.top.color = border_color


    ''' override bottom border with the specified color
    '''
    def override_bottom_border(self, border_color, forced=False):
        if border_color:
            if self.bottom is None:
                self.bottom = Border(border_dict=None)
            elif forced:
                self.bottom.color = border_color


    ''' override left border with the specified color
    '''
    def override_left_border(self, border_color):
        if border_color:
            if self.left is None:
                self.left = Border(border_dict=None)


    ''' override right border with the specified color
    '''
    def override_right_border(self, border_color):
        if border_color:
            if self.right is None:
                self.right = Border(border_dict=None)


    ''' top border
    '''
    def borders_to_latex_t(self, color_dict):
        t = None
        if self.top:
            t = self.top.border_to_latex(color_dict=color_dict)

        return t


    ''' bottom border
    '''
    def borders_to_latex_b(self, color_dict):
        b = None
        if self.bottom:
            b = self.bottom.border_to_latex(color_dict=color_dict)

        return b


    ''' left border
    '''
    def borders_to_latex_l(self, color_dict):
        l = None
        if self.left:
            l = self.left.border_to_latex(color_dict=color_dict)

        return l


    ''' right border
    '''
    def borders_to_latex_r(self, color_dict):
        r = None
        if self.right:
            r = self.right.border_to_latex(color_dict=color_dict)

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
            self.color = RgbColor(rgb_dict=border_dict.get('color'))

            # TODO: handle double
            self.style = GSHEET_LATEX_BORDER_MAPPING.get(self.style, 'solid')


    ''' string representation
    '''
    def __repr__(self):
        return f"style: {self.style}, width: {self.width}, color: {self.color}"


    ''' border
    '''
    def border_to_latex(self, color_dict):
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
        self.inches = row_height_in_inches(self.pixel_size)



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
            self.format = TextFormat(text_format_dict=new_format)
        else:
            self.start_index = None
            self.format = None


    ''' generates the latex code
    '''
    def text_format_run_to_latex(self, block_id, text, color_dict, document_footnotes, footnote_list, verbatim=False):
        latex = self.format.text_format_to_latex(block_id=block_id, text=text[self.start_index:], color_dict=color_dict, document_footnotes=document_footnotes, footnote_list=footnote_list, verbatim=verbatim)

        return latex



''' gsheet cell notes object wrapper
'''
class CellNote(object):

    ''' constructor
    '''
    def __init__(self, note_json=None, nesting_level=0):
        self.nesting_level = nesting_level
        self.free_content = False
        self.table_spacing = True
        self.page_number = False
        self.header_rows = 0

        self.style = None
        self.new_page = False
        self.keep_with_next = False
        self.keep_with_previous = False
        self.keep_line_breaks = False

        self.outline_level = 0
        self.footnotes = {}

        if note_json:
            try:
                note_dict = json.loads(note_json)
            except json.JSONDecodeError:
                note_dict = {}

            self.header_rows = int(note_dict.get('repeat-rows', 0))
            self.new_page = note_dict.get('new-page') is not None
            self.keep_with_next = note_dict.get('keep-with-next') is not None
            self.keep_with_previous = note_dict.get('keep-with-previous') is not None
            self.keep_line_breaks = note_dict.get('keep-line-breaks') is not None
            self.page_number = note_dict.get('page-number') is not None
            self.footnotes = note_dict.get('footnote')

            # content
            content = note_dict.get('content')
            if content is not None and content in ['free', 'out-of-cell']:
                self.free_content = True

            # table-spacing
            spacing = note_dict.get('table-spacing')
            if spacing is not None and spacing == 'no-spacing':
                self.table_spacing = False

            # style
            self.style = note_dict.get('style')
            if self.style is not None:
                outline_level_object = HEADING_TO_LEVEL.get(self.style, None)
                if outline_level_object:
                    self.outline_level = outline_level_object['outline-level'] + self.nesting_level
                    self.style = LEVEL_TO_HEADING[self.outline_level]

                # if style is any Title/Heading or Table or Figure, apply keep-with-next
                if self.style in LEVEL_TO_HEADING or self.style in ['Table', 'Figure']:
                    self.keep_with_next = True

            # footnotes
            if self.footnotes:
                if not isinstance(self.footnotes, dict):
                    self.footnotes = {}
                    warn(f".... found footnotes, but it is not a valid dictionary")



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



''' gsheet wrapping object wrapper
'''
class Wrapping(object):

    ''' constructor
    '''
    def __init__(self, wrap=None):
        if wrap:
            self.wrapping = WRAP_STRATEGY_MAP.get(wrap, 'WRAP')
        else:
            self.wrapping = WRAP_STRATEGY_MAP.get('WRAP')



''' Helper for cell span specification
'''
class MultiSpan(object):
    No = 'No'
    FirstCell = 'FirstCell'
    InnerCell = 'InnerCell'
    LastCell = 'LastCell'



#   ----------------------------------------------------------------------------------------------------------------
#   processors for content-types
#   ----------------------------------------------------------------------------------------------------------------

''' Table processor
'''
def process_table(section_data, config, color_dict, document_footnotes):
    # for embedded gsheets, 'contents' does not contain the actual content to render, rather we get a list of sections where each section contains the actual content
    # if section_data['contents'] is not None and 'sections' in section_data['contents']:
    #     for section in section_data['contents']['sections']:
    #         content_type = section['content-type']

    #         # force table formatter for gsheet content
    #         if content_type == 'gsheet':
    #             content_type = 'table'

    #         func = getattr(self, f"process_{content_type}")
    #         section_lines = func(section)
    #         self.document_lines = self.document_lines + section_lines

    # else:
    #     latex_section = LatexTableSection(section_data, self._config)
    #     section_lines = latex_section.section_to_latex(self.color_dict)

    latex_section = LatexTableSection(section_data=section_data, config=config)
    section_lines = latex_section.section_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)

    return section_lines


''' Gsheet processor
'''
def process_gsheet(section_data, config, color_dict, document_footnotes):
    latex_section = LatexGsheetSection(section_data=section_data, config=config)
    section_lines = latex_section.section_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)

    return section_lines


''' Table of Content processor
'''
def process_toc(section_data, config, color_dict, document_footnotes):
    latex_section = LatexToCSection(section_data=section_data, config=config)
    section_lines = latex_section.section_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)

    return section_lines


''' List of Figure processor
'''
def process_lof(section_data, config, color_dict, document_footnotes):
    latex_section = LatexLoFSection(section_data=section_data, config=config)
    section_lines = latex_section.section_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)

    return section_lines


''' List of Table processor
'''
def process_lot(section_data, config, color_dict, document_footnotes):
    latex_section = LatexLoTSection(section_data=section_data, config=config)
    section_lines = latex_section.section_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)

    return section_lines


''' pdf processor
'''
def process_pdf(section_data, config, color_dict, document_footnotes):
    latex_section = LatexPdfSection(section_data=section_data, config=config)
    section_lines = latex_section.section_to_latex(color_dict=color_dict, document_footnotes=document_footnotes)

    return section_lines


''' odt processor
'''
def process_odt(section_data, config, color_dict, document_footnotes):
    warn(f"content type [odt] not supported")

    return []


''' docx processor
'''
def process_doc(self, section_data, config, color_dict):
    warn(f"content type [docx] not supported")

    return []
