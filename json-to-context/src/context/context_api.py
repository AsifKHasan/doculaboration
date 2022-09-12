#!/usr/bin/env python3

import json
import importlib
import inspect
from pprint import pprint
 
from context.context_util import *
from helper.logger import *

#   ----------------------------------------------------------------------------------------------------------------
#   ConTeXt section objects wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' ConTeXt section base object
'''
class ContextSectionBase(object):

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

        # self.section_id = f"s_{self.section_index_text}"
        if self._section_data['link'] != '':
            self.section_id = f"s__{self.section_index_text}__{self._section_data['link']}"
        else:
            self.section_id = f"s__{self.section_index_text}__{self._section_data['content-type']}"

        self._section_data['landscape'] = 'landscape' if self._section_data['landscape'] else 'portrait'


        # master-page name
        self.landscape = self._section_data['landscape']

        self.page_spec = self._config['page-specs']['page-spec'][self._section_data['page-spec']]
        self.margin_spec = self._config['page-specs']['margin-spec'][self._section_data['margin-spec']]
        self._section_data['width'] = float(self.page_spec['width']) - float(self.margin_spec['left']) - float(self.margin_spec['right']) - float(self.margin_spec['gutter'])
        self._section_data['height'] = float(self.page_spec['height']) - float(self.margin_spec['top']) - float(self.margin_spec['bottom'])

        self.section_width = self._section_data['width']
        self.section_height = self._section_data['height']

        self.page_layout_name = self._section_data['page-layout']

        # headers and footers
        self.header_first = ContextPageHeaderFooter(content_data=section_data['header-first'], section_width=self.section_width, section_index=self.section_index, section_id=self.section_id, header_footer='header', odd_even='first', nesting_level=self.nesting_level)
        self.header_odd   = ContextPageHeaderFooter(content_data=section_data['header-odd'],   section_width=self.section_width, section_index=self.section_index, section_id=self.section_id, header_footer='header', odd_even='odd',   nesting_level=self.nesting_level)
        self.header_even  = ContextPageHeaderFooter(content_data=section_data['header-even'],  section_width=self.section_width, section_index=self.section_index, section_id=self.section_id, header_footer='header', odd_even='even',  nesting_level=self.nesting_level)
        self.footer_first = ContextPageHeaderFooter(content_data=section_data['footer-first'], section_width=self.section_width, section_index=self.section_index, section_id=self.section_id, header_footer='footer', odd_even='first', nesting_level=self.nesting_level)
        self.footer_odd   = ContextPageHeaderFooter(content_data=section_data['footer-odd'],   section_width=self.section_width, section_index=self.section_index, section_id=self.section_id, header_footer='footer', odd_even='odd',   nesting_level=self.nesting_level)
        self.footer_even  = ContextPageHeaderFooter(content_data=section_data['footer-even'],  section_width=self.section_width, section_index=self.section_index, section_id=self.section_id, header_footer='footer', odd_even='even',  nesting_level=self.nesting_level)

        self.section_contents = ContextContent(content_data=section_data.get('contents'), content_width=self.section_width, section_index=self.section_index, section_id=self.section_id, nesting_level=self.nesting_level)



    ''' generates section heading
    '''
    def get_heading(self):
        if not self.hide_heading:
            heading_title = self.heading
            if self.section != '':
                heading_title = f"{self.section} {heading_title}".strip()

            heading_title = tex_escape(heading_title)

            # headings are styles based on level
            outline_level = self.level + self.nesting_level

        else:
            heading_title = None
            outline_level = -1


        return heading_title, outline_level



    ''' get header/fotter contents
    '''
    def get_header_footer(self, color_dict, document_footnotes):
        hf_lines = []

        # get headers and footers
        if self.header_first.has_content:
            hf_lines = hf_lines +  list(map(lambda x: f"\t{x}", self.header_first.content_to_context(color_dict=color_dict, document_footnotes=document_footnotes)))

        if self.header_odd.has_content:
            hf_lines = hf_lines +  list(map(lambda x: f"\t{x}", self.header_odd.content_to_context(color_dict=color_dict, document_footnotes=document_footnotes)))

        if self.header_even.has_content:
            hf_lines = hf_lines +  list(map(lambda x: f"\t{x}", self.header_even.content_to_context(color_dict=color_dict, document_footnotes=document_footnotes)))

        if self.footer_first.has_content:
            hf_lines = hf_lines +  list(map(lambda x: f"\t{x}", self.footer_first.content_to_context(color_dict=color_dict, document_footnotes=document_footnotes)))

        if self.footer_odd.has_content:
            hf_lines = hf_lines +  list(map(lambda x: f"\t{x}", self.footer_odd.content_to_context(color_dict=color_dict, document_footnotes=document_footnotes)))

        if self.footer_even.has_content:
            hf_lines = hf_lines +  list(map(lambda x: f"\t{x}", self.footer_even.content_to_context(color_dict=color_dict, document_footnotes=document_footnotes)))


        hf_lines = mark_as_context(lines=hf_lines)

        return hf_lines



    ''' generates the ConTeXt code
    '''
    def section_to_context(self, color_dict, document_footnotes):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        section_lines = []

        # page layout changes only when a new section starts or if it is the first section
        if self.section_break or self.first_section:
            section_lines.append(f"\\setuppapersize[{self._section_data['page-spec']},{self.landscape}]")
            section_lines.append(f"\\setuplayout[{self.page_layout_name}]")


        # TODO: header/footer lines
        hf_lines = self.get_header_footer(color_dict=color_dict, document_footnotes=document_footnotes)


        # the section may have a page-break
        if self.page_break:
            section_lines.append(f"\t\\page")


        return section_lines



''' ConTeXt table section object
'''
class ContextTableSection(ContextSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        super().__init__(section_data=section_data, config=config)

 
    ''' generates the ConTeXt code
    '''
    def section_to_context(self, color_dict, document_footnotes):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        section_lines = super().section_to_context(color_dict=color_dict, document_footnotes=document_footnotes)

        # title/chapter/section/subsection/subsubsection etc.
        heading_title, outline_level = self.get_heading()
        if heading_title:
            heading_lines = ['']
            heading_lines.append(f"% {LEVEL_TO_TITLE[outline_level]} [{heading_title}]")
            heading_lines.append(f"\\{LEVEL_TO_TITLE[outline_level]}{context_option(title=heading_title)}")
            section_lines =  section_lines + heading_lines

        # get the contents
        section_lines = section_lines + self.section_contents.content_to_context(color_dict=color_dict, document_footnotes=document_footnotes)


        # wrap section in BEGIN/end comments
        section_lines = wrap_with_comment(lines=section_lines, object_type='ContextSection', object_id=self.section_id, indent_level=1)


        return [''] + section_lines



''' ConTeXt gsheet section object
'''
class ContextGsheetSection(ContextSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        super().__init__(section_data=section_data, config=config)


    ''' generates the odt code
    '''
    def section_to_context(self, color_dict, document_footnotes):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        section_lines = super().section_to_context(color_dict=color_dict, document_footnotes=document_footnotes)

        # for embedded gsheets, 'contents' does not contain the actual content to render, rather we get a list of sections where each section contains the actual content
        if self._section_data['contents'] is not None and 'sections' in self._section_data['contents']:
            # these are child contents, we need to assign indexes so that they do not overlap with parent indexes
            nesting_level = self.nesting_level + 1

            first_section = False
            # section_index = self.section_index * 100
            section_index = 0
            for section in self._section_data['contents']['sections']:
                section['nesting-level'] = nesting_level
                section['parent-section-index-text'] = self.section_index_text
                if section['section'] != '':
                    info(msg=f"writing : {section['section'].strip()} {section['heading'].strip()}", nesting_level=nesting_level)
                else:
                    info(msg=f"writing : {section['heading'].strip()}", nesting_level=nesting_level)

                section['first-section'] = True if first_section else False
                section['section-index'] = section_index

                module = importlib.import_module("context.context_api")
                func = getattr(module, f"process_{section['content-type']}")
                section_lines = section_lines + func(section_data=section, config=self._config, color_dict=color_dict, document_footnotes=document_footnotes)

                first_section = False
                section_index = section_index + 1

        section_end_lines = [f"\n% END   ContextSection: [{self.section_id}]"]

        return section_lines + section_end_lines



''' ConTeXt ToC section object
'''
class ContextToCSection(ContextSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data=section_data, config=config)


    ''' generates the ConTeXt code
    '''
    def section_to_context(self, color_dict, document_footnotes):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        section_lines = super().section_to_context(color_dict=color_dict, document_footnotes=document_footnotes)

        # table-of-contents
        content_lines = []
        content_lines.append("\\placecontent")

        # wrap in start/stop title
        content_lines = indent_and_wrap(lines=content_lines, wrap_in='title', param_string='title=Table of Contents', indent_level=1)

        # merge contents in section
        section_lines = section_lines + content_lines + ['']

        # wrap section in BEGIN/end comments
        section_lines = wrap_with_comment(lines=section_lines, object_type='ContextSection', object_id=self.section_id, indent_level=1)

        section_lines = [''] + section_lines

        return section_lines



''' ConTeXt LoT section object
'''
class ContextLoTSection(ContextSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data=section_data, config=config)


    ''' generates the ConTeXt code
    '''
    def section_to_context(self, color_dict, document_footnotes):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        section_lines = super().section_to_context(color_dict=color_dict, document_footnotes=document_footnotes)

        # table-of-contents
        content_lines = []
        content_lines.append("\\placelistoftables")

        # wrap in start/stop title
        content_lines = indent_and_wrap(lines=content_lines, wrap_in='title', param_string='title=List of Tables', indent_level=1)

        # merge contents in section
        section_lines = section_lines + content_lines + ['']

        # wrap section in BEGIN/end comments
        section_lines = wrap_with_comment(lines=section_lines, object_type='ContextSection', object_id=self.section_id, indent_level=1)

        section_lines = [''] + section_lines

        return section_lines



''' ConTeXt LoF section object
'''
class ContextLoFSection(ContextSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data=section_data, config=config)


    ''' generates the ConTeXt code
    '''
    def section_to_context(self, color_dict, document_footnotes):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        section_lines = super().section_to_context(color_dict=color_dict, document_footnotes=document_footnotes)

        # table-of-contents
        content_lines = []
        content_lines.append("\\placelistoffigures")

        # wrap in start/stop title
        content_lines = indent_and_wrap(lines=content_lines, wrap_in='title', param_string='title=List of Figures', indent_level=1)

        # merge contents in section
        section_lines = section_lines + content_lines + ['']

        # wrap section in BEGIN/end comments
        section_lines = wrap_with_comment(lines=section_lines, object_type='ContextSection', object_id=self.section_id, indent_level=1)

        section_lines = [''] + section_lines

        return section_lines



''' ConTeXt section content base object
'''
class ContextContent(object):

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

        # content_data must have 'properties' and 'data'
        if content_data and 'properties' in content_data and 'data' in content_data:
            self.has_content = True
            sheet_properties = content_data.get('properties')
            self.title = sheet_properties.get('title')

            if 'gridProperties' in sheet_properties:
                self.row_count = max(int(sheet_properties['gridProperties'].get('rowCount', 0)) - 2, 0)
                self.column_count = max(int(sheet_properties['gridProperties'].get('columnCount', 0)) - 1, 0)

                data_list = content_data.get('data')
                if isinstance(data_list, list) and len(data_list) > 0:
                    data = data_list[0]
                    self.start_row = 2
                    self.start_column = 1

                    # rowMetadata
                    for row_metadata in data.get('rowMetadata', []):
                        self.row_metadata_list.append(RowMetadata(row_metadata_dict=row_metadata))

                    # columnMetadata
                    for column_metadata in data.get('columnMetadata', []):
                        self.column_metadata_list.append(ColumnMetadata(column_metadata_dict=column_metadata))

                    # merges
                    if 'merges' in content_data:
                        for merge in content_data.get('merges', []):
                            self.merge_list.append(Merge(gsheet_merge_dict=merge, start_row=self.start_row, start_column=self.start_column))

                    # column width needs adjustment as \tabcolsep is COLSEPin. This means each column has a COLSEP inch on left and right as space which needs to be removed from column width
                    all_column_widths_in_pixel = sum(x.pixel_size for x in self.column_metadata_list[1:])
                    self.column_widths = [ (x.pixel_size * self.content_width / all_column_widths_in_pixel) - (COLSEP * 2) for x in self.column_metadata_list[1:] ]

                    # rowData
                    r = 2
                    row_data_list = data.get('rowData', [])
                    if len(row_data_list) > 2:
                        for row_data in row_data_list[2:]:
                            row = Row(section_index=self.section_index, section_id=self.section_id, row_num=r, row_data=row_data, default_format=self.default_format, section_width=self.content_width, column_widths=self.column_widths, row_height=self.row_metadata_list[r].inches, nesting_level=self.nesting_level)
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

            if first_row < 0:
                continue

            if first_col < 0:
                continue

            if first_cell is None:
                warn(f"cell [{first_cell.cell_name}] starts a span, but it is not there")
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
                        warn(f"..cell [{next_cell_in_row.cell_name}] is not empty, it must be part of another column/row merge which is an issue")

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
                    table = ContextTable(cell_matrix=self.cell_matrix, start_row=next_table_starts_in_row, end_row=(r - 1), column_widths=self.column_widths, section_index=self.section_index, section_id=self.section_id)
                    self.content_list.append(table)

                block = ContextParagraph(data_row=data_row, row_number=r, section_index=self.section_index, section_id=self.section_id)
                self.content_list.append(block)

                next_table_starts_in_row = r + 1

            # the row may start with a note of repeat-rows which means that a new table is atarting
            elif data_row.is_table_start():
                # there may be a pending/running table
                if r > next_table_starts_in_row:
                    table = ContextTable(cell_matrix=self.cell_matrix, start_row=next_table_starts_in_row, end_row=(r - 1), column_widths=self.column_widths, section_index=self.section_index, section_id=self.section_id)
                    self.content_list.append(table)

                    next_table_starts_in_row = r

            else:
                next_table_ends_in_row = r

        # there may be a pending/running table
        if next_table_ends_in_row >= next_table_starts_in_row:
            table = ContextTable(cell_matrix=self.cell_matrix, start_row=next_table_starts_in_row, end_row=next_table_ends_in_row, column_widths=self.column_widths, section_index=self.section_index, section_id=self.section_id)
            self.content_list.append(table)


    ''' generates the ConTeXt code
    '''
    def content_to_context(self, color_dict, document_footnotes):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        context_lines = []

        # iterate through tables and blocks contents
        for block in self.content_list:
            context_lines = context_lines + block.block_to_context(longtable=True, color_dict=color_dict, document_footnotes=document_footnotes)


        return context_lines



''' Page Header/Footer
'''
class ContextPageHeaderFooter(ContextContent):

    ''' constructor
        header_footer : header/footer
        odd_even      : first/odd/even
    '''
    def __init__(self, content_data, section_width, section_index, section_id, header_footer, odd_even, nesting_level):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        super().__init__(content_data=content_data, content_width=section_width, section_index=section_index, section_id=section_id, nesting_level=nesting_level)
        self.header_footer, self.odd_even = header_footer, odd_even
        self.page_header_footer_id = f"{self.header_footer}{self.odd_even}{COLUMNS[self.section_index]}"


    ''' generates the ConTeXt code
    '''
    def content_to_context(self, color_dict, document_footnotes):
        context_lines = []

        context_lines.append(f"\\providecommand\\{self.page_header_footer_id}{{}}")
        context_lines.append(f"\\renewcommand\\{self.page_header_footer_id}{{%")

        # iterate through tables and blocks contents
        first_block = True
        for block in self.content_list:
            block_lines = block.block_to_context(longtable=False, color_dict=color_dict, document_footnotes=document_footnotes, strip_comments=True, header_footer=self.header_footer)
            # block_lines = list(map(lambda x: f"\t{x}", block_lines))
            context_lines = context_lines + block_lines

            first_block = False

        context_lines.append(f"\t}}")

        return [f"% ContextPageHeaderFooter: [{self.page_header_footer_id}]"] + context_lines



''' ConTeXt Block object wrapper base class (plain ConTeXt, table, header etc.)
'''
class ContextBlock(object):

    ''' constructor
    '''
    def __init__(self, section_index, section_id):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        self.section_index, self.section_id = section_index, section_id



    ''' generates ConTeXt code
    '''
    def block_to_context(self, longtable, color_dict, document_footnotes, strip_comments=False, header_footer=None):
        pass



''' ConTeXt Table object wrapper
'''
class ContextTable(ContextBlock):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, cell_matrix, start_row, end_row, column_widths):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        super().__init__(section_index, section_id)

        self.start_row, self.end_row, self.column_widths = start_row, end_row, column_widths
        self.table_cell_matrix = cell_matrix[start_row:end_row+1]
        self.row_count = len(self.table_cell_matrix)
        col_count = len(column_widths)
        start_col = 1
        end_col = col_count

        self.table_name = f"{COLUMNS[start_col]}{self.start_row+3}-{COLUMNS[start_col]}{self.end_row+3}"
        self.table_id = f"{self.section_id}__t__{self.table_name}__{self.row_count}"

        # header row if any
        self.header_row_count = self.table_cell_matrix[0].get_cell(c=0).note.header_rows


    ''' generates the ConTeXt code
    '''
    def block_to_context(self, longtable, color_dict, document_footnotes, strip_comments=False, header_footer=None):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        # DEBUG
        # r = 1
        # print(f"Table : [{self.table_id}]")
        # for row in self.table_cell_matrix:
        #     print(f".. Row : [{row.row_id}] - {len(row.cells)} cells")
        #     r = r + 1


        # for storing the footnotes (if any) for this block
        if (self.table_id not in document_footnotes):
            document_footnotes[self.table_id] = []

        # table setups
        setup_lines = []
        loffset = f"{CONTEXT_BORDER_WIDTH_FACTOR * 3}pt"
        roffset = f"{CONTEXT_BORDER_WIDTH_FACTOR * 3}pt"
        toffset = f"{CONTEXT_BORDER_WIDTH_FACTOR * 3}pt"
        boffset = f"{CONTEXT_BORDER_WIDTH_FACTOR * 3}pt"
        rulethickness = '0.00pt'
        columndistance = f"{COLSEP}in"
        spaceinbetween = f"{ROWSEP}in"
        setup_lines.append(f"\\setupTABLE{context_option(split='repeat', header='repeat', frame='off', columndistance=columndistance, spaceinbetween=spaceinbetween, loffset=loffset, roffset=roffset, toffset=toffset, boffset=boffset, rulethickness=rulethickness)}")
        c = 1
        for col_width in self.column_widths:
            col_width = f"{col_width}in"
            setup_lines.append(f"\\setupTABLE[c][{c}]{context_option(width=col_width)}")
            c = c + 1

        # add a dummy column
        setup_lines.append(f"\\setupTABLE[c][{c}]{context_option(width='0.0in')}")

        setup_lines.append('')


        # generate row setups
        r = 1
        for row in self.table_cell_matrix:
            row_setup_lines = row.row_setups_to_context(r=r, color_dict=color_dict)
            setup_lines = setup_lines + row_setup_lines
            r = r + 1


        # generate cell setups
        r = 1
        for row in self.table_cell_matrix:
            cell_setup_lines = row.cell_setups_to_context(r=r, color_dict=color_dict)
            setup_lines = setup_lines + cell_setup_lines
            r = r + 1

        # wrap in start/stop setups
        param_string = context_option(self.table_id)
        setup_lines = indent_and_wrap(lines=setup_lines, wrap_in='setups', param_string=param_string, wrap_prefix_start='start', wrap_prefix_stop='stop', indent_level=1)


        # repeat-rows should go under TABLEhead
        table_header_lines = []
        if self.header_row_count > 0:

            # generate cell values
            for row in self.table_cell_matrix[0:self.header_row_count]:
                row_lines = row.row_contents_to_context(block_id=self.table_id, include_formatting=False, color_dict=color_dict, document_footnotes=document_footnotes, strip_comments=strip_comments, header_footer=header_footer)

                # wrap in b/e TR
                row_lines = indent_and_wrap(lines=row_lines, wrap_in='TR', wrap_prefix_start='b', wrap_prefix_stop='e', indent_level=1)

                # wrap in BEGIN/end comments
                row_lines = wrap_with_comment(lines=row_lines, object_type='Row', object_id=row.row_id, indent_level=1)
                row_lines = [''] + row_lines

                table_header_lines = table_header_lines + row_lines

            # wrap in b/e TABLEhead
            table_header_lines = indent_and_wrap(lines=table_header_lines, wrap_in='TABLEhead', wrap_prefix_start='b', wrap_prefix_stop='e', indent_level=1)


        # 
        # generate bTABLEbody cell values
        table_body_lines = []
        for row in self.table_cell_matrix[self.header_row_count:]:
            row_lines = row.row_contents_to_context(block_id=self.table_id, include_formatting=False, color_dict=color_dict, document_footnotes=document_footnotes, strip_comments=strip_comments, header_footer=header_footer)

            # wrap in b/e TR
            row_lines = indent_and_wrap(lines=row_lines, wrap_in='TR', wrap_prefix_start='b', wrap_prefix_stop='e', indent_level=1)

            # wrap in BEGIN/end comments
            row_lines = wrap_with_comment(lines=row_lines, object_type='Row', object_id=row.row_id, indent_level=1)
            row_lines = [''] + row_lines

            table_body_lines = table_body_lines + row_lines


        # FIX: we need a dummy row for resolving issue with ConTeXt
        num_columns = len(self.table_cell_matrix[0].cells) + 1
        dummy_row_lines = []
        for i in range(0, num_columns):
            # wrap in b/e TD
            dummy_row_lines = dummy_row_lines + indent_and_wrap(lines=[], wrap_in='TD', wrap_prefix_start='b', wrap_prefix_stop='e', indent_level=1)

        # wrap in b/e TR
        param_string = f"[height={DUMMY_ROW_HEIGHT}in]"
        dummy_row_lines = indent_and_wrap(lines=dummy_row_lines, wrap_in='TR', param_string=param_string, wrap_prefix_start='b', wrap_prefix_stop='e', indent_level=1)

        # wrap in BEGIN/end comments
        dummy_row_lines = wrap_with_comment(lines=dummy_row_lines, object_type='Row', object_id='dummy', indent_level=1)
        table_body_lines = table_body_lines + [''] + dummy_row_lines + ['']

        # wrap in b/e TABLEbody
        table_body_lines = indent_and_wrap(lines=table_body_lines, wrap_in='TABLEbody', wrap_prefix_start='b', wrap_prefix_stop='e', indent_level=1)


        # actual table
        table_lines = table_header_lines + table_body_lines

        # ConTeXt footnotes
        # footnote_texts = document_footnotes[self.table_id]
        # if len(footnote_texts):
        #     # append footnotetexts 
        #     table_lines.append(f"")
        #     for footnote_text_dict in footnote_texts:
        #         footnote_text = f"\t\\footnotetext[{footnote_text_dict['mark']}]{{{footnote_text_dict['text']}}}"
        #         table_lines.append(footnote_text)

        #     TODO: \\setfnsymbol for the footnotes, this needs to go before the table
        #     table_lines = [f"\\setfnsymbol{{{self.table_id}_symbols}}"] + table_lines

        
        # wrap in b/e TABLE
        param_string = context_option(setups=self.table_id)
        table_lines = indent_and_wrap(lines=table_lines, wrap_in='TABLE', param_string=param_string, wrap_prefix_start='b', wrap_prefix_stop='e', indent_level=1)

        # wrap in start/stop postponingnotes
        table_lines = indent_and_wrap(lines=table_lines, wrap_in='postponingnotes', indent_level=1)

        # wrap in BEGIN/end comments
        all_lines = setup_lines + table_lines
        all_lines = wrap_with_comment(lines=all_lines, object_type='ContextTable', object_id=self.table_id, indent_level=1)

        return [''] + all_lines



''' ConTeXt Block object wrapper
'''
class ContextParagraph(ContextBlock):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, data_row, row_number):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        super().__init__(section_index, section_id)

        self.data_row = data_row
        self.row_number = row_number
        self.paragraph_id = f"{self.section_id}__p__{self.row_number+3}"


    ''' generates the ConTeXt code
    '''
    def block_to_context(self, longtable, color_dict, document_footnotes):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        # for storing the footnotes (if any) for this block
        if (self.paragraph_id not in document_footnotes):
            document_footnotes[self.paragraph_id] = []

        block_lines = []

        # generate the block, only the first cell of the data_row to be produced
        if len(self.data_row.cells) > 0:
            row_text = self.data_row.get_cell(c=0).cell_content_to_context(block_id=self.paragraph_id, include_formatting=True, color_dict=color_dict, document_footnotes=document_footnotes)
            row_lines = list(map(lambda x: f"\t{x}", row_text))
            block_lines = block_lines + row_lines

        # ConTeXt footnotes
        footnote_texts = document_footnotes[self.paragraph_id]
        if len(footnote_texts):
            # append footnotetexts 
            block_lines.append(f"")
            for footnote_text_dict in footnote_texts:
                footnote_text = f"\t\\footnotetext[{footnote_text_dict['mark']}]{{{footnote_text_dict['text']}}}"
                block_lines.append(footnote_text)

            block_lines.append(f"")

            # \\setfnsymbol for the footnotes, this needs to go before the table
            block_lines = [f"\\setfnsymbol{{{self.id}_symbols}}"] + block_lines

        # wrap in start/stop postponingnotes
        block_lines = indent_and_wrap(lines=block_lines, wrap_in='postponingnotes', indent_level=1)
        block_lines = []

        # wrap in BEGIN/end comments
        block_lines = wrap_with_comment(lines=block_lines, object_type='ContextParagraph', object_id=self.paragraph_id, indent_level=1)

        return [''] + block_lines



#   ----------------------------------------------------------------------------------------------------------------
#   gsheet cell wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' gsheet Cell object wrapper
'''
class Cell(object):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, row_num, col_num, value, default_format, column_widths, row_height, nesting_level):
        self.section_index, self.section_id, self.row_num, self.col_num, self.column_widths, self.default_format, self.nesting_level = section_index, section_id, row_num, col_num, column_widths, default_format, nesting_level
        self.cell_id = f"{COLUMNS[self.col_num+1]}{self.row_num+1}"
        self.cell_name = self.cell_id
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
                self.cell_value = ContentValue(section_index=self.section_index, section_id=self.section_id, effective_format=self.effective_format, content_value=self.value['contents'])

            elif 'userEnteredValue' in self.value:
                if 'image' in self.value['userEnteredValue']:
                    self.cell_value = ImageValue(section_index=self.section_index, section_id=self.section_id, effective_format=self.effective_format, image_value=self.value['userEnteredValue']['image'])

                else:
                    if len(self.text_format_runs):
                        self.cell_value = TextRunValue(section_index=self.section_index, section_id=self.section_id, effective_format=self.effective_format, text_format_runs=self.text_format_runs, formatted_value=self.formatted_value)

                    elif self.note.page_numbering:
                        self.cell_value = PageNumberValue(section_index=self.section_index, section_id=self.section_id, effective_format=self.effective_format, page_numbering=self.note.page_numbering)

                    else:
                        if self.note.script and self.note.script == 'latex':
                            self.cell_value = ContextValue(section_index=self.section_index, section_id=self.section_id, effective_format=self.effective_format, string_value=self.value['userEnteredValue'], formatted_value=self.formatted_value, nesting_level=self.nesting_level, outline_level=self.note.outline_level)
                        
                        else:
                            self.cell_value = StringValue(section_index=self.section_index, section_id=self.section_id, effective_format=self.effective_format, string_value=self.value['userEnteredValue'], formatted_value=self.formatted_value, nesting_level=self.nesting_level, outline_level=self.note.outline_level)

            else:
                self.cell_value = StringValue(section_index=self.section_index, section_id=self.section_id, effective_format=self.effective_format, string_value='', formatted_value=self.formatted_value, nesting_level=self.nesting_level, outline_level=self.note.outline_level)

        else:
            # value can have a special case it can be an empty ditionary when the cell is an inner cell of a column merge
            self.merge_spec.multi_col = MultiSpan.No
            self.note = CellNote()
            self.cell_value = None
            self.formatted_value = ''
            self.effective_format = None
            self.user_entered_format = None
            self.is_empty = True


    ''' string representation
    '''
    def __repr__(self):
        s = f"{self.cell_name:>4}, value: {not self.is_empty:<1}, mr: {self.merge_spec.multi_row:<9}, mc: {self.merge_spec.multi_col:<9} [{self.formatted_value}]"
        return s


    ''' ConTeXt code for cell content
    '''
    def cell_content_to_context(self, block_id, include_formatting, color_dict, document_footnotes):
        content_lines = []

        # the content is not valid for multirow LastCell and InnerCell
        if self.merge_spec.multi_row in [MultiSpan.No, MultiSpan.FirstCell] and self.merge_spec.multi_col in [MultiSpan.No, MultiSpan.FirstCell]:
            if self.cell_value:

                # get the ConTeXt code
                cell_value = self.cell_value.value_to_context(block_id=block_id, container_width=self.cell_width, container_height=self.cell_height, color_dict=color_dict, document_footnotes=document_footnotes, footnote_list=self.note.footnotes)

                # handle new-page defined in notes
                if self.note.new_page:
                    content_lines.append(f"\\page")

                # handle keep-with-previous defined in notes
                if self.note.keep_with_previous:
                    content_lines.append(f"\\page[no]")

                if self.note.style:
                    # handle styles defined in notes
                    if self.note.style == 'Figure':
                        # caption for figure
                        content_lines.append(f"% Figure heading")

                    elif self.note.style == 'Table':
                        # caption for table
                        content_lines.append(f"% Table heading")

                    elif self.note.style:
                        # some custom style needs to be applied
                        heading_tag = CONTEXT_HEADING_MAP.get(self.note.style)
                        if heading_tag:
                            content_lines.append(f"% {heading_tag}")

                        else:
                            warn(f"style : {self.note.style} not defined")

                else:
                    content_lines.append(cell_value)

            else:
                warn(f"NO VALUE : {self.cell_id}")

        return content_lines


    ''' ConTeXt code for cell setups
    '''
    def cell_setups_to_context(self, r, color_dict):
        cell_setup_lines = []

        color_dict[self.effective_format.bgcolor.key()] = self.effective_format.bgcolor.value()
        if not self.is_empty:
            cell_format_lines = self.effective_format.cellformat_to_context_options(color_dict=color_dict)
            cell_setup_lines.append(f"% Cell setups [{self.cell_id}]")
            for line in cell_format_lines:
                cell_setup_lines.append(f"\\setupTABLE[{self.col_num+1}][{r}]{line}")
    
            cell_setup_lines.append('')

        return cell_setup_lines


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



''' gsheet Row object wrapper
'''
class Row(object):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, row_num, row_data, default_format, section_width, column_widths, row_height, nesting_level):
        self.section_index, self.section_id, self.row_num, self.default_format, self.section_width, self.column_widths, self.row_height, self.nesting_level = section_index, section_id, row_num, default_format, section_width, column_widths, row_height, nesting_level
        self.row_id = f"{self.row_num+1}"
        self.row_name = f"row: [{self.row_id}]"

        self.cells = []
        c = 0
        values = row_data.get('values', [])
        if len(values) > 1:
            for value in values[1:]:
                cell = Cell(section_index=self.section_index, section_id=self.section_id, row_num=self.row_num, col_num=c, value=value, default_format=self.default_format, column_widths=self.column_widths, row_height=self.row_height, nesting_level=self.nesting_level)
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


    ''' generates the ConTeXt code for row setups
    '''
    def row_setups_to_context(self, r, color_dict):
        row_setup_lines = []
        row_setup_options = context_option(minheight=f"{self.row_height}in")
        row_setup_lines.append(f"% Row setups [{self.row_id}]")
        row_setup_lines.append(f"\\setupTABLE[r][{r}]{row_setup_options}")
        row_setup_lines.append('')

        return row_setup_lines


    ''' generates the ConTeXt code for cell setups
    '''
    def cell_setups_to_context(self, r, color_dict):
        cell_setup_lines = []
        for cell in self.cells:
            if cell is not None:
                if not cell.is_empty:
                    cell_setup_lines = cell_setup_lines + cell.cell_setups_to_context(r=r, color_dict=color_dict)

        return cell_setup_lines


    ''' generates the ConTeXt code
    '''
    def row_contents_to_context(self, block_id, include_formatting, color_dict, document_footnotes, strip_comments=False, header_footer=None):
        row_lines = []

        # gets the cell ConTeXt lines
        c = 0
        produce_cell = True
        for cell in self.cells:
            cell_lines = []
            if cell is None:
                warn(f"{self.row_name} has a Null cell at {c}")
                produce_cell = False

            else:
                cell_lines = cell.cell_content_to_context(block_id=block_id, include_formatting=include_formatting, color_dict=color_dict, document_footnotes=document_footnotes)

            if produce_cell and len(cell_lines):
                # wrap in b/e TR, param_string is the marge spec
                cell_lines = indent_and_wrap(lines=cell_lines, wrap_in='TD', param_string=cell.merge_spec.to_context_option(), wrap_prefix_start='b', wrap_prefix_stop='e', indent_level=1)

                # wrap in BEGIN/end comments
                cell_lines = wrap_with_comment(lines=cell_lines, object_type='Cell', object_id=cell.cell_id, begin_suffix=cell.merge_spec.to_string(), indent_level=1)
                cell_lines = [''] + cell_lines


                row_lines = row_lines + cell_lines

            c = c + 1

        # FIX: add a dummy cell
        # wrap in b/e TD
        dummy_cell_lines = indent_and_wrap(lines=[], wrap_in='TD', wrap_prefix_start='b', wrap_prefix_stop='e', indent_level=1)

        # wrap in BEGIN/end comments
        dummy_cell_lines = wrap_with_comment(lines=dummy_cell_lines, object_type='Cell', object_id='dummy', indent_level=1)
        row_lines = row_lines + [''] + dummy_cell_lines + ['']

        return row_lines



''' gsheet cell value object wrapper
'''
class CellValue(object):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, effective_format, nesting_level=0, outline_level=0):
        self.effective_format = effective_format
        self.nesting_level = nesting_level
        self.outline_level = outline_level
        self.section_index = section_index
        self.section_id = section_id



''' string type CellValue
'''
class StringValue(CellValue):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, effective_format, string_value, formatted_value, nesting_level=0, outline_level=0):
        super().__init__(section_index=section_index, section_id=section_id, effective_format=effective_format, nesting_level=nesting_level, outline_level=outline_level)
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


    ''' generates the ConTeXt code
    '''
    def value_to_context(self, block_id, container_width, container_height, color_dict, document_footnotes, footnote_list):
        verbatim = False
        context_code = self.effective_format.text_format.text_format_to_context(block_id=block_id, text=self.value, color_dict=color_dict, document_footnotes=document_footnotes, footnote_list=footnote_list, verbatim=verbatim)
        # context_code = f"{{{self.value}}}"

        return context_code



''' LaTex type CellValue
'''
class ContextValue(CellValue):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, effective_format, string_value, formatted_value, nesting_level=0, outline_level=0):
        super().__init__(section_index=section_index, section_id=section_id, effective_format=effective_format, nesting_level=nesting_level, outline_level=outline_level)
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


    ''' generates the ConTeXt code
    '''
    def value_to_context(self, block_id, container_width, container_height, color_dict, document_footnotes, footnote_list):
        verbatim = True
        # context_code = self.effective_format.text_format.text_format_to_context(block_id=block_id, text=self.value, color_dict=color_dict, document_footnotes=document_footnotes, footnote_list=footnote_list, verbatim=verbatim)
        context_code = f"{{{self.value}}}"

        return context_code



''' text-run type CellValue
'''
class TextRunValue(CellValue):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, effective_format, text_format_runs, formatted_value, nesting_level=0, outline_level=0):
        super().__init__(section_index=section_index, section_id=section_id, effective_format=effective_format, nesting_level=nesting_level, outline_level=outline_level)
        self.text_format_runs = text_format_runs
        self.formatted_value = formatted_value


    ''' string representation
    '''
    def __repr__(self):
        s = f"text-run : [{self.formatted_value}]"
        return s


    ''' generates the ConTeXt code
    '''
    def value_to_context(self, block_id, container_width, container_height, color_dict, document_footnotes, footnote_list):
        verbatim = False
        run_value_list = []
        processed_idx = len(self.formatted_value)
        for text_format_run in reversed(self.text_format_runs):
            text = self.formatted_value[:processed_idx]
            run_value_list.insert(0, text_format_run.text_format_run_to_context(block_id=block_id, text=text, color_dict=color_dict, document_footnotes=document_footnotes, footnote_list=footnote_list, verbatim=verbatim))
            processed_idx = text_format_run.start_index

        return ''.join(run_value_list)



''' page-number type CellValue
'''
class PageNumberValue(CellValue):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, effective_format, page_numbering='long', nesting_level=0, outline_level=0):
        super().__init__(section_index=section_index, section_id=section_id, effective_format=effective_format, nesting_level=nesting_level, outline_level=outline_level)
        self.page_numbering = page_numbering


    ''' string representation
    '''
    def __repr__(self):
        s = f"page-number"
        return s


    ''' generates the ConTeXt code
    '''
    def value_to_context(self, block_id, container_width, container_height, color_dict, document_footnotes, footnote_list):
        if self.page_numbering == 'short':
            context_code = "\\thepage"
        else:
            context_code = "\\thepage\\ of \\pageref{LastPage}"

        return context_code



''' image type CellValue
'''
class ImageValue(CellValue):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, effective_format, image_value, nesting_level=0, outline_level=0):
        super().__init__(section_index=section_index, section_id=section_id, effective_format=effective_format, nesting_level=nesting_level, outline_level=outline_level)
        self.value = image_value


    ''' string representation
    '''
    def __repr__(self):
        s = f"image : [{self.value}]"
        return s


    ''' generates the ConTeXt code
    '''
    def value_to_context(self, block_id, container_width, container_height, color_dict, document_footnotes, footnote_list):
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

        image_options = context_option(width=f"{image_width_in_inches}in", height=f"{image_height_in_inches}in")
        context_code = f"\\externalfigure[{os_specific_path(self.value['path'])}]{image_options}"

        # image horizontal alignment
        image_halign = self.effective_format.halign.image_halign()
        image_valign = self.effective_format.valign.image_valign()

        if image_halign:
            context_code = f"\\{image_halign}{{{context_code}}}"

        if image_valign:
            context_code = f"\\{image_valign}{{{context_code}}}"


        return context_code



''' content type CellValue
'''
class ContentValue(CellValue):

    ''' constructor
    '''
    def __init__(self, section_index, section_id, effective_format, content_value, nesting_level=0, outline_level=0):
        super().__init__(section_index=section_index, section_id=section_id, effective_format=effective_format, nesting_level=nesting_level, outline_level=outline_level)
        self.value = content_value


    ''' string representation
    '''
    def __repr__(self):
        s = f"content : [{self.value['sheets'][0]['properties']['title']}]"
        return s


    ''' generates the ConTeXt code
    '''
    def value_to_context(self, block_id, container_width, container_height, color_dict, document_footnotes, footnote_list):
        section_contents = ContextContent(content_data=self.value, content_width=container_width, section_index=self.section_index, section_id=self.section_id, nesting_level=self.nesting_level)
        return section_contents.content_to_context(color_dict=color_dict, document_footnotes=document_footnotes)



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


    ''' CellFormat to ConTeXt setups
    '''
    def cellformat_to_context_options(self, color_dict):
        cell_setup_options = []

        # formatting
        background = 'color'
        backgroundcolor = self.bgcolor
        foregroundcolor = self.text_format.fgcolor
        align = f"{{{self.halign.cell_halign()},{self.valign.cell_valign()}}}"
        cellformat_options = context_option(background=background, backgroundcolor=backgroundcolor, foregroundcolor=foregroundcolor, align=align)
        cell_setup_options.append(cellformat_options)

        cell_setup_options = cell_setup_options + self.borders.borders_to_context_options(color_dict)

        return cell_setup_options



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


    ''' generate ConTeXt code
    '''
    def text_format_to_context(self, block_id, text, color_dict, document_footnotes, footnote_list, verbatim=False):
        color_dict[self.fgcolor.key()] = self.fgcolor.value()

        # process inline blocks (footnotes, LaTeX, etc. (if any))
        content = f"{process_inline_blocks(block_id=block_id, text_content=text, document_footnotes=document_footnotes, footnote_list=footnote_list, verbatim=verbatim)}"

        style = None
        if self.is_bold:
            if self.is_italic:
                style = f"\\bi"

            else:
                style = f"\\bf"
        else:
            if self.is_italic:
                style = f"\\it"

        if style:
            content = f"{style}{{{content}}}"

        if self.is_underline:
            content = f"\\underbar{{{content}}}"

        if self.is_strikethrough:
            content = f"\\overstrike{{{content}}}"


        # color, font, font-size
        font_spec = ''
        if self.font_family != '':
            font_spec = f"\\globalfsize{{{self.font_size}pt}}\\switchtobodyfont[{self.font_family}]\\color[{self.fgcolor.key()}]"
        else:
            font_spec = f"\\globalfsize{{{self.font_size}pt}}\\color[{self.fgcolor.key()}]"

        context_code = f"{font_spec}{{{content}}}"

        return context_code



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


    ''' borders as cConTeXt options
    '''
    def borders_to_context_options(self, color_dict):
        option_lines = []
        if self.top:
            self.top.border_to_context(color_dict=color_dict)
            option_line = context_option(topframe='on', framecolor=self.top.color.key(), rulethickness=f"{self.top.width}pt")
            option_lines.append(option_line)

        if self.bottom:
            self.bottom.border_to_context(color_dict=color_dict)
            option_line = context_option(bottomframe='on', framecolor=self.bottom.color.key(), rulethickness=f"{self.bottom.width}pt")
            option_lines.append(option_line)

        if self.left:
            self.left.border_to_context(color_dict=color_dict)
            option_line = context_option(leftframe='on', framecolor=self.left.color.key(), rulethickness=f"{self.left.width}pt")
            option_lines.append(option_line)

        if self.right:
            self.right.border_to_context(color_dict=color_dict)
            option_line = context_option(rightframe='on', framecolor=self.right.color.key(), rulethickness=f"{self.right.width}pt")
            option_lines.append(option_line)

        return option_lines



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
            self.width = int(border_dict.get('width')) * CONTEXT_BORDER_WIDTH_FACTOR
            self.color = RgbColor(rgb_dict=border_dict.get('color'))

            # TODO: handle double
            self.style = GSHEET_LATEX_BORDER_MAPPING.get(self.style, 'solid')


    ''' string representation
    '''
    def __repr__(self):
        return f"style: {self.style}, width: {self.width}, color: {self.color}"


    ''' border
    '''
    def border_to_context(self, color_dict):
        color_dict[self.color.key()] = self.color.value()



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


    ''' return something like [nr=n, nc=m]
    '''
    def to_context_option(self):
        if self.col_span > 1:
            if self.row_span > 1:
                return context_option(nr=self.row_span, nc=self.col_span)
            else:
                return context_option(nc=self.col_span)

        else:
            if self.row_span > 1:
                return context_option(nr=self.row_span)
            else:
                return None

        return None



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


    ''' string representation
    '''
    def __repr__(self):
        return f"{COLUMNS[self.start_col]}{self.start_row+1}:{COLUMNS[self.end_col-1]}{self.end_row}"



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


    ''' generates the ConTeXt code
    '''
    def text_format_run_to_context(self, block_id, text, color_dict, document_footnotes, footnote_list, verbatim=False):
        context_code = self.format.text_format_to_context(block_id=block_id, text=text[self.start_index:], color_dict=color_dict, document_footnotes=document_footnotes, footnote_list=footnote_list, verbatim=verbatim)
        # context_code = text=text[self.start_index:]

        return context_code



''' gsheet cell notes object wrapper
'''
class CellNote(object):

    ''' constructor
    '''
    def __init__(self, note_json=None, nesting_level=0):
        self.nesting_level = nesting_level
        self.free_content = False
        self.table_spacing = True
        self.page_numbering = None
        self.header_rows = 0

        self.style = None
        self.new_page = False
        self.keep_with_next = False
        self.keep_with_previous = False
        self.keep_line_breaks = False

        self.script = None

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
            self.footnotes = note_dict.get('footnote')

            # content
            content = note_dict.get('content')
            if content is not None and content in ['free', 'out-of-cell']:
                self.free_content = True

            # script
            script = note_dict.get('script')
            if script is not None and script in ['latex']:
                self.script = script

            # page-number
            page_numbering = note_dict.get('page-number')
            if page_numbering is not None and page_numbering in ['long', 'short']:
                self.page_numbering = page_numbering

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
        self.valign = valign


    ''' cell vertical alignment for ConTeXt
    '''
    def cell_valign(self):
        if self.valign:
            return CELL_VALIGN_MAP.get(self.valign)
        else:
            return CELL_VALIGN_MAP.get('TOP')


    ''' image vertical alignment for ConTeXt
    '''
    def image_valign(self):
        if self.valign:
            return IMAGE_POSITION.get(self.valign)
        else:
            return IMAGE_POSITION.get('middle')



''' gsheet horizontal alignment object wrapper
'''
class HorizontalAlignment(object):

    ''' constructor
    '''
    def __init__(self, halign=None):
        self.halign = halign


    ''' cell horizontal alignment for ConTeXt
    '''
    def cell_halign(self):
        if self.halign:
            return CELL_HALIGN_MAP.get(self.halign)
        else:
            return CELL_HALIGN_MAP.get('LEFT')


    ''' image horizontal alignment for ConTeXt
    '''
    def image_halign(self):
        if self.halign:
            return IMAGE_POSITION.get(self.halign)
        else:
            return IMAGE_POSITION.get('center')



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
    context_section = ContextTableSection(section_data=section_data, config=config)
    section_lines = context_section.section_to_context(color_dict=color_dict, document_footnotes=document_footnotes)

    return section_lines


''' Gsheet processor
'''
def process_gsheet(section_data, config, color_dict, document_footnotes):
    context_section = ContextGsheetSection(section_data=section_data, config=config)
    section_lines = context_section.section_to_context(color_dict=color_dict, document_footnotes=document_footnotes)

    return section_lines


''' Table of Content processor
'''
def process_toc(section_data, config, color_dict, document_footnotes):
    context_section = ContextToCSection(section_data=section_data, config=config)
    section_lines = context_section.section_to_context(color_dict=color_dict, document_footnotes=document_footnotes)

    return section_lines


''' List of Figure processor
'''
def process_lof(section_data, config, color_dict, document_footnotes):
    context_section = ContextLoFSection(section_data=section_data, config=config)
    section_lines = context_section.section_to_context(color_dict=color_dict, document_footnotes=document_footnotes)

    return section_lines


''' List of Table processor
'''
def process_lot(section_data, config, color_dict, document_footnotes):
    context_section = ContextLoTSection(section_data=section_data, config=config)
    section_lines = context_section.section_to_context(color_dict=color_dict, document_footnotes=document_footnotes)

    return section_lines


''' pdf processor
'''
def process_pdf(section_data, config, color_dict, document_footnotes):
    context_section = ContextPdfSection(section_data=section_data, config=config)
    section_lines = context_section.section_to_context(color_dict=color_dict, document_footnotes=document_footnotes)

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
