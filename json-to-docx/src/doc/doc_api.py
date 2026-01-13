#!/usr/bin/env python

import json
import importlib
import inspect
from pprint import pprint

from doc.doc_util import *
from helper.logger import *

#   ----------------------------------------------------------------------------------------------------------------
#   docx (not oo section, gsheet section) objects wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' Docx section base object
'''
class DocxSectionBase(object):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")


        self._config = config
        self._section_data = section_data
        self._doc = self._config['docx']

        self.section_meta = self._section_data['section-meta']
        self.section_prop = self._section_data['section-prop']

        self.label = self.section_prop['label']
        self.heading = self.section_prop['heading']
        self.heading_style = self.section_prop['heading-style']
        self.level = self.section_prop['level']
        self.section_break = self.section_prop['section-break']
        self.page_break = self.section_prop['page-break']
        self.hide_heading = self.section_prop['hide-heading']

        self.landscape = self.section_prop['landscape']

        self.page_spec_name = self.section_prop['page-spec']
        self.page_spec = self._config['page-specs']['page-spec'][self.page_spec_name]

        self.margin_spec_name = self.section_prop['margin-spec']
        self.margin_spec = self._config['page-specs']['margin-spec'][self.margin_spec_name]

        self.autocrop = self.section_prop['autocrop']
        self.page_bg = self.section_prop['page-bg']

        self.bookmark = self.section_prop['bookmark']
        self.background_image = self.section_prop['background-image']

        self.document_index = self.section_meta['document-index']
        self.document_name = self.section_meta['document-name']
        self.section_index = self.section_meta['section-index']
        self.section_name = self.section_meta['section-name']
        self.orientation = self.section_meta['orientation']
        self.first_section = self.section_meta['first-section']
        self.document_first_section = self.section_meta['document-first-section']
        self.different_firstpage = self.section_meta['different-firstpage']
        self.different_odd_even_pages = self.section_meta['different-odd-even-pages']
        self.nesting_level = self.section_meta['nesting-level']
        self.page_layout_name = self.section_meta['page-layout']

        self.section_id = f"D{str(self.document_index).zfill(3)}--{self.document_name}__S{str(self.section_index).zfill(3)}--{self.section_name}"


        if self.landscape:
            self.page_width = float(self.page_spec['height'])
            self.page_height = float(self.page_spec['width'])
            self.section_width = self.page_width - float(self.margin_spec['left']) - float(self.margin_spec['right']) - float(self.margin_spec['gutter'])
            self.section_height = self.page_height - float(self.margin_spec['top']) - float(self.margin_spec['bottom'])
        else:
            self.page_width = float(self.page_spec['width'])
            self.page_height = float(self.page_spec['height'])
            self.section_width = self.page_width - float(self.margin_spec['left']) - float(self.margin_spec['right']) - float(self.margin_spec['gutter'])
            self.section_height = self.page_height - float(self.margin_spec['top']) - float(self.margin_spec['bottom'])


        # create or get the docx section
        self.docx_section, new_section = add_or_update_document_section(doc=self._doc, page_spec=self.page_spec, margin_spec=self.margin_spec, orientation=self.orientation, different_firstpage=self.different_firstpage, section_break=self.section_break, page_break=self.page_break, first_section=self.first_section, different_odd_even_pages=self.different_odd_even_pages, background_image_path=self.background_image)


        # headers and footers
        if new_section:
            self.header_first = DocxPageHeaderFooter(doc=self._doc, content_data=self._section_data['header-first'], section_width=self.section_width, section_index=self.section_index, header_footer='header', odd_even='first', nesting_level=self.nesting_level)
            self.header_odd   = DocxPageHeaderFooter(doc=self._doc, content_data=self._section_data['header-odd'],   section_width=self.section_width, section_index=self.section_index, header_footer='header', odd_even='odd',   nesting_level=self.nesting_level)
            self.header_even  = DocxPageHeaderFooter(doc=self._doc, content_data=self._section_data['header-even'],  section_width=self.section_width, section_index=self.section_index, header_footer='header', odd_even='even',  nesting_level=self.nesting_level)
            self.footer_first = DocxPageHeaderFooter(doc=self._doc, content_data=self._section_data['footer-first'], section_width=self.section_width, section_index=self.section_index, header_footer='footer', odd_even='first', nesting_level=self.nesting_level)
            self.footer_odd   = DocxPageHeaderFooter(doc=self._doc, content_data=self._section_data['footer-odd'],   section_width=self.section_width, section_index=self.section_index, header_footer='footer', odd_even='odd',   nesting_level=self.nesting_level)
            self.footer_even  = DocxPageHeaderFooter(doc=self._doc, content_data=self._section_data['footer-even'],  section_width=self.section_width, section_index=self.section_index, header_footer='footer', odd_even='even',  nesting_level=self.nesting_level)

        else:
            self.header_first = None
            self.header_odd   = None
            self.header_even  = None
            self.footer_first = None
            self.footer_odd   = None
            self.footer_even  = None


        self.section_contents = DocxContent(doc=self._doc, content_data=section_data.get('contents'), content_width=self.section_width, nesting_level=self.nesting_level)


    ''' get section heading
    '''
    def get_heading(self):
        # identify what style the heading will be and its content
        if not self.hide_heading:
            heading_title = self.heading
            if self.label != '':
                heading_title = f"{self.label} {heading_title}".strip()

            # headings are styles based on level
            outline_level = self.level + self.nesting_level

            if outline_level == 0:
                style_name = 'Title'
            else:
                style_name = f"Heading{outline_level}"

        else:
            heading_title = None
            style_name = 'Normal'
            outline_level = 0


        return heading_title, outline_level, style_name


    ''' Header/Footer processing
    '''
    def process_header_footer(self):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        if self.header_first:
            self.header_first.content_to_doc(container=self.docx_section.first_page_header)

        if self.header_odd:
            self.header_odd.content_to_doc(container=self.docx_section.header)

        if self.header_even:
            self.header_even.content_to_doc(container=self.docx_section.even_page_header)

        if self.footer_first:
            self.footer_first.content_to_doc(container=self.docx_section.first_page_footer)

        if self.footer_odd:
            self.footer_odd.content_to_doc(container=self.docx_section.footer)

        if self.footer_even:
            self.footer_even.content_to_doc(container=self.docx_section.even_page_footer)


    ''' generates the docx code
    '''
    def section_to_doc(self):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        self.process_header_footer()

        heading_text, outline_level, style_name = self.get_heading()
        if heading_text:
            paragraph = create_paragraph(doc=self._doc, container=self._doc, paragraph_attributes={'stylename': style_name}, text_content=heading_text, outline_level=outline_level, bookmark=self.bookmark)

            if self.heading_style:
                if self.heading_style not in self._config['custom-styles']:
                    warn(f"custom style [{self.heading_style}] not defined in style-specs")

                else:
                    trace(f"applying custom style [{self.heading_style}] to heading")
                    apply_custom_style(doc=self._doc, style_spec=self._config['custom-styles'][self.heading_style], paragraph=paragraph, nesting_level=self.nesting_level+1)



''' Docx table section object
'''
class DocxTableSection(DocxSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        super().__init__(section_data, config)


    ''' generates the docx code
    '''
    def section_to_doc(self):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        super().section_to_doc()
        self.section_contents.content_to_doc(container=self._doc)


''' Docx gsheet section object
'''
class DocxGsheetSection(DocxSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data, config)


    ''' generates the docx code
    '''
    def section_to_doc(self):
        super().section_to_doc()

        # for embedded gsheets, 'contents' does not contain the actual content to render, rather we get a list of sections where each section contains the actual content
        if self._section_data['contents'] is not None and 'sections' in self._section_data['contents']:
            # process the sections
            section_list_to_doc(self._section_data['contents']['sections'], self._config)


''' Docx ToC section object
'''
class DocxToCSection(DocxSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data, config)


    ''' generates the docx code
    '''
    def section_to_doc(self):
        super().section_to_doc()
        create_index(self._doc, index_type='toc')


''' Docx LoT section object
'''
class DocxLoTSection(DocxSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data, config)


    ''' generates the docx code
    '''
    def section_to_doc(self):
        super().section_to_doc()
        create_index(self._doc, index_type='lot')


''' Docx LoF section object
'''
class DocxLoFSection(DocxSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data, config)


    ''' generates the docx code
    '''
    def section_to_doc(self):
        super().section_to_doc()
        create_index(self._doc, index_type='lof')


''' Docx Pdf section object
'''
class DocxPdfSection(DocxSectionBase):

    ''' constructor
    '''
    def __init__(self, section_data, config):
        super().__init__(section_data, config)


    ''' generates the docx code
    '''
    def section_to_doc(self):
        super().section_to_doc()

        # the images go one after another
        text_attributes = {'fontsize': 2}
        style_attributes = {}
        text_format_attributes = {'textalign': TEXT_HALIGN_MAP['CENTER']}
        
        if 'contents' in self._section_data:
            if self._section_data['contents'] and 'images' in self._section_data['contents']:
                for i, image in enumerate(self._section_data['contents']['images']):
                    paragraph_attributes = {'breakbefore': 'page'}

                    if self.page_bg == True:
                        paragraph = self._doc.add_paragraph()
                        apply_paragraph_attributes(paragraph=paragraph, paragraph_attributes=paragraph_attributes)
                        create_page_background(doc=self._doc, background_image_path=image['path'], page_width_inches=self.page_width, page_height_inches=self.page_height)

                    else:
                        paragraph = self._doc.add_paragraph()
                        apply_paragraph_attributes(paragraph=paragraph, paragraph_attributes=paragraph_attributes)
                        format_container(container=paragraph, attributes=text_format_attributes, it_is_a_table_cell=False)

                        image_width_in_inches, image_height_in_inches = image['width'], image['height']
                        fit_within_width = self.section_width
                        fit_within_height = self.section_height
                        fit_within_height = fit_within_height - PDF_PAGE_HEIGHT_OFFSET

                        # print(f"image  size [{image_width_in_inches}x{image_height_in_inches}] aspect ratio [{image_height_in_inches/image_width_in_inches}]")
                        # print(f"target size [{fit_within_width}x{fit_within_height}] aspect ratio [{fit_within_height/fit_within_width}]")
                        image_width_in_inches, image_height_in_inches = fit_width_height(fit_within_width=fit_within_width, fit_within_height=fit_within_height, width_to_fit=image_width_in_inches, height_to_fit=image_height_in_inches)
                        # print(f"final  size [{image_width_in_inches}x{image_height_in_inches}] aspect ratio [{image_height_in_inches/image_width_in_inches}]")
                        # print()
                        
                        insert_image(container=paragraph, picture_path=image['path'], width=image_width_in_inches, height=image_height_in_inches)
                        # insert_image(container=paragraph, picture_path=image['path'], height=image_height_in_inches)


''' Docx section content base object
'''
class DocxContent(object):

    ''' constructor
    '''
    def __init__(self, doc, content_data, content_width, nesting_level):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")
        self._doc = doc
        self.content_data = content_data
        self.content_width = content_width
        self.nesting_level = nesting_level

        self.title = None
        self.row_count = 0
        self.column_count = 0

        self.start_row = 0
        self.start_column = 0

        self.cell_matrix = []
        self.row_metadata_list = []
        self.column_metadata_list = []
        self.merge_list = []

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
                        self.row_metadata_list.append(RowMetadata(row_metadata))

                    # columnMetadata
                    for column_metadata in data.get('columnMetadata', []):
                        self.column_metadata_list.append(ColumnMetadata(column_metadata))

                    # merges
                    if 'merges' in content_data:
                        for merge in content_data.get('merges', []):
                            self.merge_list.append(Merge(merge, self.start_row, self.start_column))

                    all_column_widths_in_pixel = sum(x.pixel_size for x in self.column_metadata_list[1:])
                    self.column_widths = [ (x.pixel_size * self.content_width / all_column_widths_in_pixel) for x in self.column_metadata_list[1:] ]

                    # rowData
                    r = 2
                    row_data_list = data.get('rowData', [])
                    if len(row_data_list) > 2:
                        for row_data in row_data_list[2:]:
                            new_row = Row(
                                doc=self._doc, 
                                row_num=r,
                                row_data=row_data,
                                section_width=self.content_width,
                                column_widths=self.column_widths,
                                row_height=self.row_metadata_list[r].inches,
                                nesting_level=self.nesting_level,
                            )
                            self.cell_matrix.append(new_row)
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
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

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

            if first_row < 0:
                continue

            if first_col < 0:
                continue

            if first_cell is None:
                warn(f"cell [{first_cell.cell_name}] starts a span, but it is not there")
                continue

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

            # considering merges, we have effective cell width and height
            first_cell.effective_cell_width = sum(first_cell.column_widths[first_cell.col_num:first_cell.col_num + col_span])
            effective_row_height = 0
            for r in range(first_row, last_row):
                effective_row_height = effective_row_height + self.cell_matrix[r].row_height

            first_cell.effective_cell_height = effective_row_height

            # for row spans, subsequent cells in the same column of the FirstCell will be either empty or missing, iterate through the next rows
            for r in range(first_row, last_row):
                next_row_object = self.cell_matrix[r]
                row_height = next_row_object.row_height

                # we may have empty cells in this same row which are part of this column merge, we need to mark their multi_col property correctly
                for c in range(first_col, last_col):
                    # exclude the very first cell
                    if r == first_row and c == first_col:
                        continue

                    next_cell_in_row = next_row_object.get_cell(c)

                    if next_cell_in_row is None:
                        # the cell may not be existing at all, we have to create
                        next_cell_in_row = Cell(doc=self._doc, row_num=r+2, col_num=c, value=None, column_widths=first_cell.column_widths, row_height=row_height, nesting_level=self.nesting_level)
                        next_row_object.insert_cell(c, next_cell_in_row)

                    if next_cell_in_row.is_empty:
                        # the cell is a newly inserted one, its format should be the same (for borders, colors) as the first cell so that we can draw borders properly
                        # debug(f"cell [{next_cell_in_row.cell_id}] is empty")
                        next_cell_in_row.copy_format_from(first_cell)

                        # mark cells for multicol only if it is multicol
                        if col_span > 1:
                            if c == first_col:
                                # the last cell of the merge to be marked as LastCell
                                next_cell_in_row.merge_spec.multi_col = MultiSpan.FirstCell

                            elif c == last_col-1:
                                # the last cell of the merge to be marked as LastCell
                                next_cell_in_row.merge_spec.multi_col = MultiSpan.LastCell

                            else:
                                # the inner cells of the merge to be marked as InnerCell
                                next_cell_in_row.merge_spec.multi_col = MultiSpan.InnerCell

                        else:
                            next_cell_in_row.merge_spec.multi_col = MultiSpan.No

                        # mark cells for multirow only if it is multirow
                        if row_span > 1:
                            if r == first_row:
                                # the last cell of the merge to be marked as LastCell
                                next_cell_in_row.merge_spec.multi_row = MultiSpan.FirstCell

                            elif r == last_row-1:
                                # the last cell of the merge to be marked as LastCell
                                next_cell_in_row.merge_spec.multi_row = MultiSpan.LastCell

                            else:
                                # the inner cells of the merge to be marked as InnerCell
                                next_cell_in_row.merge_spec.multi_row = MultiSpan.InnerCell

                        else:
                            next_cell_in_row.merge_spec.multi_row = MultiSpan.No

                    else:
                        warn(f"..cell [{next_cell_in_row.cell_name}] is not empty, it must be part of another column/row merge which is an issue")

    ''' processes the cells to split the cells into tables and blocks and orders the tables and blocks properly
    '''
    def split(self):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        # we have a concept of in-cell content and out-of-cell (free) content
        # in-cell contents are treated as part of a table, while out-of-cell (free) contents are treated as independent paragraphs, images etc. (blocks)
        next_table_starts_in_row = 0
        next_table_ends_in_row = 0
        for r in range(0, self.row_count):
            if len(self.cell_matrix) <= r:
                continue

            # the first cell of the row tells us whether it is in-cell or out-of-cell
            data_row = self.cell_matrix[r]

            # do extra processing on rows
            # data_row.preprocess_row()

            if data_row.is_free_content():
                # there may be a pending/running table
                if r > next_table_starts_in_row:
                    table = DocxTable(self.cell_matrix, next_table_starts_in_row, r - 1, self.column_widths)
                    self.content_list.append(table)

                block = DocxParagraph(data_row, r)
                self.content_list.append(block)

                next_table_starts_in_row = r + 1

            # the row may start with a note of repeat-rows which means that a new table is atarting
            elif data_row.is_table_start():
                # there may be a pending/running table
                if r > next_table_starts_in_row:
                    table = DocxTable(self.cell_matrix, next_table_starts_in_row, r - 1, self.column_widths)
                    self.content_list.append(table)

                    next_table_starts_in_row = r

            else:
                next_table_ends_in_row = r

        # there may be a pending/running table
        if next_table_ends_in_row >= next_table_starts_in_row:
            table = DocxTable(self.cell_matrix, next_table_starts_in_row, next_table_ends_in_row, self.column_widths)
            self.content_list.append(table)


    ''' generates the docx code
        container may be doc, header/footer or a Cell
    '''
    def content_to_doc(self, container):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        # iterate through tables and blocks contents
        for block in self.content_list:
            block.block_to_doc(container=container, container_width=self.content_width)


''' Docx Page Header Footer object
'''
class DocxPageHeaderFooter(DocxContent):

    ''' constructor
        header_footer : header/footer
        odd_even      : first/odd/even(left)
    '''
    def __init__(self, doc, content_data, section_width, section_index, header_footer, odd_even, nesting_level):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        self.nesting_level = nesting_level
        super().__init__(doc=doc, content_data=content_data, content_width=section_width, nesting_level=nesting_level)
        self.header_footer, self.odd_even = header_footer, odd_even
        self.page_header_footer_id = f"{self.header_footer}-{self.odd_even}-{section_index}"


''' Docx Block object wrapper base class (plain docx, table, header etc.)
'''
class DocxBlock(object):

    ''' constructor
    '''
    def __init__(self):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")
        pass


''' Docx Table object wrapper
'''
class DocxTable(DocxBlock):

    ''' constructor
    '''
    def __init__(self, cell_matrix, start_row, end_row, column_widths):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        self.start_row, self.end_row, self.column_widths = start_row, end_row, column_widths
        self.table_cell_matrix = cell_matrix[start_row:end_row+1]
        self.row_count = len(self.table_cell_matrix)
        self.col_count = len(column_widths)
        self.table_name = f"Table_{random_string()}"

        # header row if any
        first_cell = self.table_cell_matrix[0].get_cell(0)
        if first_cell:
            self.header_row_count = self.table_cell_matrix[0].get_cell(0).note.header_rows
        else:
            self.header_row_count = 0

    ''' string representation
    '''
    def __repr__(self):
        lines = []
        lines.append(f"[{self.col_count}x{self.row_count}] table")
        for row in self.table_cell_matrix:
            lines.append(str(row))

        return '\n'.join(lines)

    ''' generates the docx code
    '''
    def block_to_doc(self, container, container_width):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        num_cols = len(self.column_widths)
        num_rows = len(self.table_cell_matrix)

        tbl = create_table(container=container, num_rows=num_rows, num_cols=num_cols, container_width=container_width)

        # table-columns
        for c in range(0, len(self.column_widths)):
            col_width = self.column_widths[c]
            table_column = tbl.columns[c]
            table_column.width = Inches(col_width)

        # iterate rows and cells to render the table's contents
        if tbl:
            for row_index in range(0, len(self.table_cell_matrix)):
                row = self.table_cell_matrix[row_index]
                table_row = tbl.rows[row_index]
                row.row_to_doc_table_row(table=tbl, table_row=table_row)

            # header rows
            for r in reversed(range(0, self.header_row_count)):
                trace(f"repeating row : {r}")
                set_repeat_table_header(tbl.rows[r])

            # merge cells
            for row_index in range(0, len(self.table_cell_matrix)):
                row = self.table_cell_matrix[row_index]
                for col_index in range(0, len(row.cells)):
                    cell = row.cells[col_index]
                    if cell:
                        if cell.merge_spec.multi_row == MultiSpan.FirstCell or cell.merge_spec.multi_col == MultiSpan.FirstCell:
                            # this is the first cell of a merge, we need to get the last cell
                            start_table_cell = tbl.cell(row_index, col_index)
                            end_table_cell = tbl.cell(row_index + cell.merge_spec.row_span-1, col_index + cell.merge_spec.col_span-1)
                            start_table_cell.merge(end_table_cell)
                    else:
                        warn(f"[{self.__class__.__name__} : {inspect.stack()[0][3]}] - row [{row_index}], cell [{col_index}] is None, that should not be")


            #  decorate cells
            for row_index in range(0, len(self.table_cell_matrix)):
                row = self.table_cell_matrix[row_index]
                for col_index in range(0, len(row.cells)):
                    cell = row.cells[col_index]
                    if cell:
                        cell.decorate_cell()
                    else:
                        # warn(f"[{self.__class__.__name__} : {inspect.stack()[0][3]}] - row [{row_index}], cell [{col_index}] is None, that should not be")
                        pass


''' Docx Block object wrapper
'''
class DocxParagraph(DocxBlock):

    ''' constructor
    '''
    def __init__(self, data_row, row_number):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        self.data_row = data_row
        self.row_number = row_number

    ''' string representation
    '''
    def __repr__(self):
        return f"__para__"

    ''' generates the docx code
    '''
    def block_to_doc(self, container, container_width):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        # generate the block, only the first cell of the data_row to be produced
        if len(self.data_row.cells) > 0:
            # We take the first cell, the cell will take the whole row width
            cell_to_produce = self.data_row.get_cell(0)
            cell_to_produce.cell_width = sum(cell_to_produce.column_widths)

            cell_to_produce.cell_to_doc(container=container)


#   ----------------------------------------------------------------------------------------------------------------
#   gsheet cell wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' gsheet Cell object wrapper
'''
class Cell(object):

    ''' constructor
    '''
    def __init__(self, doc, row_num, col_num, value, column_widths, row_height, nesting_level):
        self._doc = doc
        self.row_num, self.col_num, self.column_widths, self.nesting_level  = row_num, col_num, column_widths, nesting_level
        self.cell_id = f"{COLUMNS[self.col_num+1]}{self.row_num+1}"
        self.cell_name = self.cell_id
        self.value = value
        self.text_format_runs = []
        self.cell_width = self.column_widths[self.col_num]
        self.cell_height = row_height

        self.merge_spec = CellMergeSpec()
        self.background = None
        self.note = None
        self.effective_format = CellFormat(format_dict=None)
        self.user_entered_format = CellFormat(format_dict=None)

        # considering merges, we have effective cell width and height
        self.effective_cell_width = self.cell_width
        self.effective_cell_height = self.cell_height

        if self.value:
            bg_dict = self.value.get('background', {})
            if bg_dict:
                bg_dict['container-width'] = self.effective_cell_width
                bg_dict['container-height'] = self.effective_cell_height
                self.background = CellBackground(bg_dict)

            note_dict = self.value.get('notes', {})
            self.note = CellNote(note_dict=note_dict)

            self.formatted_value = self.value.get('formattedValue', '')

            self.effective_format = CellFormat(self.value.get('effectiveFormat'))

            for text_format_run in self.value.get('textFormatRuns', []):
                self.text_format_runs.append(TextFormatRun(run_dict=text_format_run, default_format=self.effective_format.text_format.source))

            # presence of userEnteredFormat makes the cell non-empty
            if 'userEnteredFormat' in self.value:
                self.user_entered_format = CellFormat(format_dict=self.value.get('userEnteredFormat'))
                self.is_empty = False

                # HACK: handle background-color - if user_entered_format does not have backgroundColor, omit backgroundColor from effective_format
                if 'backgroundColor' in self.value.get('userEnteredFormat'):
                    self.effective_format.bgcolor = self.user_entered_format.bgcolor

            else:
                self.user_entered_format = None
                self.is_empty = True


            # we need to identify exactly what kind of value the cell contains
            if 'contents' in self.value:
                self.cell_value = ContentValue(doc=self._doc, effective_format=self.effective_format, content_value=self.value['contents'], background=self.background)

            elif 'userEnteredValue' in self.value:
                if 'image' in self.value['userEnteredValue']:
                    self.cell_value = ImageValue(doc=self._doc, effective_format=self.effective_format, image_value=self.value['userEnteredValue']['image'], background=self.background)

                else:
                    if len(self.text_format_runs):
                        self.cell_value = TextRunValue(doc=self._doc, effective_format=self.effective_format, text_format_runs=self.text_format_runs, formatted_value=self.formatted_value, background=self.background)

                    else:
                        if self.note.script and self.note.script == 'latex':
                            self.cell_value = LatexValue(doc=self._doc, effective_format=self.effective_format, string_value=self.value['userEnteredValue'], formatted_value=self.formatted_value, background=self.background, nesting_level=self.nesting_level, outline_level=self.note.outline_level)

                        else:
                            self.cell_value = StringValue(doc=self._doc, effective_format=self.effective_format, string_value=self.value['userEnteredValue'], formatted_value=self.formatted_value, background=self.background, nesting_level=self.nesting_level, outline_level=self.note.outline_level)

            else:
                self.cell_value = StringValue(doc=self._doc, effective_format=self.effective_format, string_value=None, formatted_value=self.formatted_value, background=self.background, nesting_level=self.nesting_level, outline_level=self.note.outline_level)

        else:
            # value can have a special case it can be an empty dictionary when the cell is an inner cell of a merge
            self.cell_value = None
            self.formatted_value = ''
            self.is_empty = True



    ''' string representation
    '''
    def __repr__(self):
        s = f".... {self.cell_name:>4}, value: {not self.is_empty:<1}, mr: {self.merge_spec.multi_row:<9}, mc: {self.merge_spec.multi_col:<9} [{self.formatted_value[0:50]}]"
        # s = f".... {self.cell_name:>4}, value: {not self.is_empty:<1}, mr: {self.merge_spec.multi_row:<9}, mc: {self.merge_spec.multi_col:<9} [{self.effective_format.borders}]"
        return s

    ''' docx code for cell content
    '''
    def cell_to_doc_table_cell(self, table, table_cell):
        # trace(f"{self}")
        self.table_cell = table_cell
        table_cell.width = Inches(self.cell_width)
        if self.effective_format:
            self.cell_to_doc(container=table_cell)

    ''' docx code for cell content
    '''
    def cell_to_doc(self, container):
        # trace(f"{self}")

        # for string and image it returns a paragraph, for embedded content a list
        # the content is not valid for multirow LastCell and InnerCell
        if self.merge_spec.multi_row in [MultiSpan.No, MultiSpan.FirstCell] and self.merge_spec.multi_col in [MultiSpan.No, MultiSpan.FirstCell]:
            if self.cell_value:
                if self.effective_format:
                    table_cell_attributes = self.effective_format.table_cell_attributes(cell_merge_spec=self.merge_spec, force_halign=self.note.force_halign, angle=self.note.angle)
                    if self.effective_format.text_format:
                        text_attributes = self.effective_format.text_format.text_attributes()
                    else:
                        text_attributes = {}
                else:
                    table_cell_attributes = {}
                    text_attributes = {}

                if self.note:
                    paragraph_attributes = self.note.paragraph_attributes()
                    footnote_list = self.note.footnotes
                else:
                    paragraph_attributes = {}
                    footnote_list = {}

                where = self.cell_value.value_to_doc(container=container, container_width=self.effective_cell_width, container_height=self.effective_cell_height, paragraph_attributes=paragraph_attributes, text_attributes=text_attributes, footnote_list=footnote_list, bookmark=self.note.bookmark)

                # do not aaply table-cell format here, it needs to be done after the merging is done
                if not is_table_cell(container) and where:
                    format_container(container=where, attributes=table_cell_attributes, it_is_a_table_cell=False)

    ''' apply formatting for table cell
    '''
    def decorate_cell(self):
        # if self.cell_value:
        if self.note:
            force_halign = self.note.force_halign
            angle = self.note.angle
        else:
            force_halign = False
            angle = 0

        table_cell_attributes = self.effective_format.table_cell_attributes(cell_merge_spec=self.merge_spec, force_halign=force_halign, angle=angle)
        format_container(container=self.table_cell, attributes=table_cell_attributes, it_is_a_table_cell=True)

    ''' Copy format from the cell passed
    '''
    def copy_format_from(self, from_cell):
        self.effective_format = from_cell.effective_format


''' gsheet Row object wrapper
'''
class Row(object):

    ''' constructor
    '''
    def __init__(self, doc, row_num, row_data, section_width, column_widths, row_height, nesting_level):
        self._doc = doc
        self.row_num, self.section_width, self.column_widths, self.row_height, self.nesting_level = row_num, section_width, column_widths, row_height, nesting_level
        self.row_name = f"row: [{self.row_num+1}]"

        self.cells = []
        c = 0
        values = row_data.get('values', [])
        if len(values) > 1:
            for value in values[1:]:
                self.cells.append(Cell(doc=self._doc, row_num=self.row_num, col_num=c, value=value, column_widths=self.column_widths, row_height=self.row_height, nesting_level=self.nesting_level))
                c = c + 1


    ''' string representation
    '''
    def __repr__(self):
        lines = []
        lines.append(f".. row [{self.row_name}]")
        for cell in self.cells:
            lines.append(str(cell))

        return '\n'.join(lines)


    ''' preprocess row
        does something automatically even if this is not in the gsheet
        1. make single cell row with style defined out-of-cell and keep-with-next
    '''
    def preprocess_row(self):
        # if the row is a single cell row (only the first cell is empty) and the cell contains a style note, make it out-of-cell and make it keep-with-next
        if len(self.cells) > 0:
            first_cell = self.cells[0]
            if not first_cell.is_empty and first_cell.note.style is not None:
                # if the other cells all are empty, we mark it out-of-cell and keep-with-next
                non_empty_cell_found = False
                for cell in self.cells[1:]:
                    if cell.is_empty == False:
                        non_empty_cell_found = True
                        break
                    else:
                        pass

                if non_empty_cell_found == False:
                    first_cell.note.free_content = True
                    first_cell.note.keep_with_next = True


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

    ''' it is true when the first cell has a free_content true value
        the first cell may be free_content when
        1. it contains a note {'content': 'out-of-cell'}
        2. it contains a note {'style': '...'} and it is the only non-empty cell in the row
    '''
    def is_free_content(self):
        if len(self.cells) > 0:
            # the first cell is the relevant cell only
            if self.cells[0]:
                if self.cells[0].note:
                    return self.cells[0].note.free_content
                else:
                    return False
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
                if self.cells[0].note:
                    return (self.cells[0].note.header_rows > 0)
                else:
                    return False
            else:
                return False
        else:
            return False

    ''' generates the docx code
    '''
    def row_to_doc_table_row(self, table, table_row):
        # trace(f"{self}")
        table_row.height = Inches(self.row_height)

        # iterate over the cells
        for cell_index in range(0, len(self.cells)):
            cell = self.cells[cell_index]
            if cell:
                cell.cell_to_doc_table_cell(table=table, table_cell=table_row.cells[cell_index])
            else:
                warn(f"[{self.__class__.__name__} : {inspect.stack()[0][3]}] - row [{self.row_num}], cell [{cell_index}] is None, that should not be")


''' gsheet text format object wrapper
'''
class TextFormat(object):

    ''' constructor
    '''
    def __init__(self, text_format_dict=None):
        self.source = text_format_dict
        if self.source:
            self.fgcolor = RgbColor(text_format_dict.get('foregroundColor'))
            if 'fontFamily' in text_format_dict:
                self.font_family = text_format_dict['fontFamily']

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


    ''' attributes dict for TextProperties
    '''
    def text_attributes(self):
        attributes = {}

        attributes['color'] = self.fgcolor
        if self.font_family:
            attributes['fontname'] = self.font_family

        attributes['fontsize'] = self.font_size

        if self.is_bold:
            attributes['fontweight'] = "bold"

        if self.is_italic:
            attributes['fontstyle'] = "italic"

        if self.is_underline:
            attributes['textunderlinestyle'] = "solid"
            attributes['textunderlinewidth'] = "auto"
            attributes['textunderlinecolor'] = "font-color"

        if self.is_strikethrough:
            attributes['textlinethroughstyle'] = "solid"
            attributes['textlinethroughtype'] = "single"

        return attributes


''' gsheet cell value object wrapper
'''
class CellValue(object):

    ''' constructor
    '''
    def __init__(self, doc, effective_format, background, nesting_level=0, outline_level=0):
        self._doc = doc
        self.effective_format = effective_format
        self.background = background
        self.nesting_level = nesting_level
        self.outline_level = outline_level


''' string type CellValue
'''
class StringValue(CellValue):

    ''' constructor
    '''
    def __init__(self, doc, effective_format, string_value, formatted_value, background, nesting_level=0, outline_level=0, directives=True):
        super().__init__(doc=doc, effective_format=effective_format, background=background, nesting_level=nesting_level, outline_level=outline_level)

        if formatted_value:
            self.value = formatted_value
        else:
            if string_value and 'stringValue' in string_value:
                self.value = string_value['stringValue']
            else:
                self.value = ''

        self.directives = directives


    ''' string representation
    '''
    def __repr__(self):
        s = f"string : [{self.value}]"
        return s


    ''' generates the docx code
    '''
    def value_to_doc(self, container, container_width, container_height, paragraph_attributes, text_attributes, footnote_list={}, bookmark={}):
        paragraph = create_paragraph(doc=self._doc, container=container, text_content=self.value, paragraph_attributes=paragraph_attributes, text_attributes=text_attributes, background=self.background, outline_level=self.outline_level, footnote_list=footnote_list, bookmark=bookmark, directives=self.directives)
        return paragraph


''' LaTex type CellValue
'''
class LatexValue(CellValue):

    ''' constructor
    '''
    def __init__(self, doc, effective_format, string_value, formatted_value, background, nesting_level=0, outline_level=0):
        super().__init__(doc=doc, effective_format=effective_format, background=background, nesting_level=nesting_level, outline_level=outline_level)
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
        s = f"latex : [{self.value}]"
        return s


    ''' generates the docx code
    '''
    def value_to_doc(self, container, container_width, container_height, paragraph_attributes, text_attributes, footnote_list={}, bookmark={}):
        paragraph = create_paragraph(doc=self._doc, container=container, paragraph_attributes=paragraph_attributes, text_attributes=text_attributes, background=self.background, outline_level=self.outline_level, bookmark=bookmark)
        create_latex(container=paragraph, latex_content=self.value)

        return paragraph


''' text-run type CellValue
'''
class TextRunValue(CellValue):

    ''' constructor
    '''
    def __init__(self, doc, effective_format, text_format_runs, formatted_value, background, nesting_level=0, outline_level=0):
        super().__init__(doc=doc, effective_format=effective_format, background=background, nesting_level=nesting_level, outline_level=outline_level)
        self.text_format_runs = text_format_runs
        self.formatted_value = formatted_value


    ''' string representation
    '''
    def __repr__(self):
        s = f"text-run : [{self.formatted_value}]"
        return s


    ''' generates the docx code
    '''
    def value_to_doc(self, container, container_width, container_height, paragraph_attributes, text_attributes, footnote_list={}, bookmark={}):
        run_value_list = []
        processed_idx = len(self.formatted_value)
        for text_format_run in reversed(self.text_format_runs):
            text = self.formatted_value[:processed_idx]
            run_value_list.insert(0, text_format_run.text_attributes(text))
            processed_idx = text_format_run.start_index

        paragraph = create_paragraph(doc=self._doc, container=container, run_list=run_value_list, background=self.background, footnote_list=footnote_list, bookmark=bookmark)
        return paragraph


''' image type CellValue
'''
class ImageValue(CellValue):

    ''' constructor
    '''
    def __init__(self, doc, effective_format, image_value, background, nesting_level=0, outline_level=0):
        super().__init__(doc=doc, effective_format=effective_format, background=background, nesting_level=nesting_level, outline_level=outline_level)
        self.value = image_value


    ''' string representation
    '''
    def __repr__(self):
        s = f"image : [{self.value}]"
        return s


    ''' generates the docx code
    '''
    def value_to_doc(self, container, container_width, container_height, paragraph_attributes, text_attributes, footnote_list={}, bookmark={}):
        # even now the width may exceed actual cell width, we need to adjust for that
        dpi_x = DPI if self.value['dpi'][0] == 0 else self.value['dpi'][0]
        dpi_y = DPI if self.value['dpi'][1] == 0 else self.value['dpi'][1]
        image_width_in_pixel = self.value['size'][0]
        image_height_in_pixel = self.value['size'][1]
        image_width_in_inches =  image_width_in_pixel / dpi_x
        image_height_in_inches = image_height_in_pixel / dpi_y

        text_attributes['fontsize'] = 0
        picture_path = self.value['path']

        if self.value['mode'] in [1, 2, 3, 4]:
            image_width_in_inches, image_height_in_inches = fit_width_height(fit_within_width=container_width, fit_within_height=container_height, width_to_fit=image_width_in_inches, height_to_fit=image_height_in_inches)

        else:
            warn(f"unknown mode [{self.value['mode']}] for image [{picture_path}]")

        where = insert_image(container=container, picture_path=picture_path, width=image_width_in_inches, height=image_height_in_inches, bookmark=bookmark)
        return where


''' content type CellValue
'''
class ContentValue(CellValue):

    ''' constructor
    '''
    def __init__(self, doc, effective_format, content_value, background, nesting_level=0, outline_level=0):
        super().__init__(doc=doc, effective_format=effective_format, background=background, nesting_level=nesting_level, outline_level=outline_level)
        self.value = content_value


    ''' string representation
    '''
    def __repr__(self):
        s = f"content : [{self.value['sheets'][0]['properties']['title']}]"
        return s


    ''' generates the docx code
    '''
    def value_to_doc(self, container, container_width, container_height, paragraph_attributes, text_attributes, footnote_list={}, bookmark={}):
        self.contents = DocxContent(doc=self._doc, content_data=self.value, content_width=container_width, nesting_level=self.nesting_level)
        self.contents.content_to_doc(container=container)
        return None


''' gsheet cell format object wrapper
'''
class CellFormat(object):

    ''' constructor
    '''
    def __init__(self, format_dict):
        self.bgcolor = None
        self.borders = None
        self.padding = None
        self.halign = None
        self.valign = None
        self.text_format = None
        self.wrapping = None
        self.text_rotation = None
        self.bgcolor_style = None

        if format_dict:
            self.bgcolor = RgbColor(format_dict.get('backgroundColor'))
            self.borders = Borders(format_dict.get('borders'))
            self.padding = Padding(format_dict.get('padding'))
            self.halign = HorizontalAlignment(format_dict.get('horizontalAlignment'))
            self.valign = VerticalAlignment(format_dict.get('verticalAlignment'))
            self.text_format = TextFormat(format_dict.get('textFormat'))
            self.wrapping = Wrapping(format_dict.get('wrapStrategy'))
            self.text_rotation = TextRotation(format_dict.get('textRotation'))

            bgcolor_style_dict = format_dict.get('backgroundColorStyle')
            if bgcolor_style_dict:
                self.bgcolor_style = RgbColor(bgcolor_style_dict.get('rgbColor'))


    ''' attributes dict for TableCellProperties
    '''
    def table_cell_attributes(self, cell_merge_spec, force_halign=False, angle=0):
        attributes = {}

        if force_halign:
            attributes['textalign'] = self.halign.halign
        else:
            if self.halign:
                attributes['textalign'] = self.halign.halign

        if self.valign:
            attributes['verticalalign'] = self.valign.valign

        if self.wrapping:
            attributes['wrapoption'] = self.wrapping.wrapping

        if self.text_rotation and self.text_rotation.angle is not None:
            attributes['angle'] = self.text_rotation.angle
            
        if angle == 90:
            attributes['angle'] = 'btLr'
        elif angle == 270:
            attributes['angle'] = 'tbRl'
            

        if self.bgcolor:
            if self.bgcolor_style:
                if not self.bgcolor_style.is_white():
                    attributes['backgroundcolor'] = self.bgcolor_style.value()
            else:
                attributes['backgroundcolor'] = self.bgcolor.value()

        more_attributes = {}
        if self.borders and self.padding:
            more_attributes = {**self.borders.table_cell_attributes(cell_merge_spec), **self.padding.table_cell_attributes()}

        return {**attributes, **more_attributes}


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


    ''' table-cell attributes
    '''
    def table_cell_attributes(self, cell_merge_spec):
        attributes = {}

        # top and bottom
        if cell_merge_spec.multi_row in [MultiSpan.No, MultiSpan.FirstCell]:
            if self.top:
                attributes['top'] = self.top.value()

            if self.bottom:
                attributes['bottom'] = self.bottom.value()

        if cell_merge_spec.multi_row in [MultiSpan.LastCell]:
            if self.bottom:
                attributes['bottom'] = self.bottom.value()


        # left and right
        if cell_merge_spec.multi_col in [MultiSpan.No, MultiSpan.FirstCell]:
            if self.left:
                attributes['start'] = self.left.value()

            if self.right:
                attributes['end'] = self.right.value()

        if cell_merge_spec.multi_col in [MultiSpan.LastCell]:
            if self.right:
                attributes['end'] = self.right.value()

        return {'borders': attributes}


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
            self.width = border_dict.get('width')
            self.color = RgbColor(border_dict.get('color'))

            # TODO: handle double
            self.style = GSHEET_OXML_BORDER_MAPPING.get(self.style, 'single')


    ''' string representation
    '''
    def __repr__(self):
        return f"{self.width}pt {self.style} {self.color.value()}"


    ''' value
    	top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
    	bottom={"sz": 12, "color": "#00FF00", "val": "single"},
    	start={"sz": 24, "val": "dashed", "shadow": "true"},
    	end={"sz": 12, "val": "dashed"},
    '''
    def value(self):
    	return {"sz": self.width * 2, "val": self.style, "color": self.color.value(), "space": "0"}


''' Cell Merge spec wrapper
'''
class CellMergeSpec(object):
    def __init__(self):
        self.multi_col = MultiSpan.No
        self.multi_row = MultiSpan.No

        self.col_span = 1
        self.row_span = 1


    ''' string representation
    '''
    def to_string(self):
        return f"multicolumn: {self.multi_col}, multirow: {self.multi_row}"


    ''' table-cell attributes
    '''
    def table_cell_attributes(self):
        attributes = {}

        if self.col_span > 1:
            attributes['numbercolumnsspanned'] = self.col_span

        if self.row_span:
            attributes['numberrowsspanned'] = self.row_span

        return attributes


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
        return f"{COLUMNS[self.start_col+1]}{self.start_row+3}:{COLUMNS[self.end_col]}{self.end_row+2}"


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
        return '#' + ''.join('{:02x}'.format(a) for a in [self.red, self.green, self.blue])


    ''' is the color white
    '''
    def is_white(self):
        if self.red == 255 and self.green == 255 and self.blue == 255:
            return True

        return False


''' gsheet cell padding object wrapper
'''
class Padding(object):

    ''' constructor
    '''
    def __init__(self, padding_dict=None):
        if padding_dict:
            # self.top = int(padding_dict.get('top', 0))
            # self.right = int(padding_dict.get('right', 0))
            # self.bottom = int(padding_dict.get('bottom', 0))
            # self.left = int(padding_dict.get('left', 0))
            self.top = 0
            self.right = 30
            self.bottom = 0
            self.left = 30
        else:
            self.top = 0
            self.right = 30
            self.bottom = 0
            self.left = 30


    ''' string representation
    '''
    def table_cell_attributes(self):
        attributes = {}

        attributes['top'] = self.top
        attributes['end'] = self.right
        attributes['bottom'] = self.bottom
        attributes['start'] = self.left

        return {'padding': attributes}


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


    ''' generates the docx code
    '''
    def text_attributes(self, text):
        return {'text': text[self.start_index:], 'text-attributes': self.format.text_attributes()}


''' gsheet cell notes object wrapper
'''
class CellNote(object):

    ''' constructor
    '''
    def __init__(self, note_dict=None, nesting_level=0):
        self.nesting_level = nesting_level

        self.outline_level = 0
        self.free_content = False
        self.content = note_dict.get('content', None)
        self.style = note_dict.get('style', 'Normal')
        self.header_rows = int(note_dict.get('repeat-rows', 0))
        self.new_page = note_dict.get('new-page', False)
        self.force_halign = note_dict.get('force-halign', False)
        self.keep_with_next = note_dict.get('keep-with-next', False)
        self.keep_with_previous = note_dict.get('keep-with-previous', False)
        self.keep_line_breaks = note_dict.get('keep-line-breaks', False)
        self.directives = note_dict.get('directives', True)
        self.angle = int(note_dict.get("angle", 0))

        self.footnotes = note_dict.get('footnote', {})
        self.bookmark = note_dict.get('bookmark', {})

        self.table_spacing = note_dict.get('table-spacing', True)
        self.script = note_dict.get('script')

        # content
        if self.content is not None and self.content in ['free', 'out-of-cell']:
            self.free_content = True

        # script
        if self.script is not None and self.script not in ['latex']:
            self.script = None

        # table-spacing
        if self.table_spacing is not None and self.table_spacing == 'no-spacing':
            self.table_spacing = False

        # style
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
                warn(f"found footnotes, but it is not a valid dictionary", nesting_level=nesting_level+1)

        # bookmark
        if self.bookmark:
            if not isinstance(self.bookmark, dict):
                self.bookmark = {}
                warn(f"found bookmark, but it is not a valid dictionary", nesting_level=nesting_level+1)


    ''' paragraph attributes dict
    '''
    def paragraph_attributes(self):
        attributes = {}

        if self.style is not None:
            attributes['stylename'] = self.style

        if self.new_page:
            attributes['breakbefore'] = 'page'

        if self.keep_with_next:
            attributes['keepwithnext'] = 'always'


        if self.keep_with_previous:
            attributes['keepwithprevious'] = 'always'

        return attributes


''' gsheet cell background wrapper
    NOTE: for now only image background is supported
'''
class CellBackground(object):

    ''' constructor
    '''
    def __init__(self, bg_dict={}):
        self.bg_dict = bg_dict
        self.file_path = bg_dict['file-path']
        self.image_width = bg_dict['image-width']
        self.image_height = bg_dict['image-height']
        self.container_width = bg_dict['container-width']
        self.container_height = bg_dict['container-height']
        self.aspect_ratio = self.image_width/self.image_height
    
    ''' attributes dict for TableCellProperties
    '''
    def table_cell_properties_attributes(self):
        attributes = {}

        return attributes


''' gsheet vertical alignment object wrapper
'''
class VerticalAlignment(object):

    ''' constructor
    '''
    def __init__(self, valign=None):
        if valign:
            self.valign = TEXT_VALIGN_MAP.get(valign)
        else:
            self.valign = TEXT_VALIGN_MAP.get('TOP')


''' gsheet horizontal alignment object wrapper
'''
class HorizontalAlignment(object):

    ''' constructor
    '''
    def __init__(self, halign=None):
        # TODO: it is overriding text style
        if halign is None or halign in ['LEFT', 'JUSTIFY', 'CENTER', 'RIGHT']:
            self.halign = TEXT_HALIGN_MAP.get(halign)
        else:
            self.halign = TEXT_HALIGN_MAP.get('LEFT')


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


''' gsheet textRotation object wrapper
'''
class TextRotation(object):

    ''' constructor
    '''
    def __init__(self, text_rotation=None):
        self.angle = None
        
        if text_rotation:
            if 'vertical' in text_rotation and text_rotation['vertical'] == False:
                self.angle = 'btLr'
                
            if 'angle' in text_rotation:
                self.angle = 'tbRl'


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
def process_table(section_data, config):
    section = DocxTableSection(section_data, config)
    section.section_to_doc()


''' Gsheet processor
'''
def process_gsheet(section_data, config):
    section = DocxGsheetSection(section_data, config)
    section.section_to_doc()


''' Table of Content processor
'''
def process_toc(section_data, config):
    section = DocxToCSection(section_data, config)
    section.section_to_doc()


''' List of Figure processor
'''
def process_lof(section_data, config):
    section = DocxLoFSection(section_data, config)
    section.section_to_doc()


''' List of Table processor
'''
def process_lot(section_data, config):
    section = DocxLoTSection(section_data, config)
    section.section_to_doc()


''' pdf processor
'''
def process_pdf(section_data, config):
    section = DocxPdfSection(section_data, config)
    section.section_to_doc()


''' odt processor
'''
def process_odt(section_data, config):
    warn(f"content type [odt] not supported")


''' docx processor
'''
def process_doc(section_data, config):
    warn(f"content type [docx] not supported")
