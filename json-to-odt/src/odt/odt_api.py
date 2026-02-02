#!/usr/bin/env python

from helper.config_service import ConfigService
from odt.odt_util import *
from helper.logger import *

#   ----------------------------------------------------------------------------------------------------------------
#   odt section (not oo section, gsheet section) objects wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' Odt section base object
'''
class OdtSectionBase(object):

    ''' constructor
    '''
    def __init__(self, odt, section_data, nesting_level=0):
        self._odt = odt
        self._section_data = section_data

        self.section_meta = self._section_data['section-meta']
        self.section_prop = self._section_data['section-prop']

        self.label = self.section_prop['label']
        self.level = self.section_prop['level']
        self.section_break = self.section_prop['section-break']
        self.page_break = self.section_prop['page-break']
        self.hide_heading = self.section_prop['hide-heading']
        self.heading = self.section_prop['heading']
        self.heading_style = self.section_prop['heading-style']

        self.landscape = self.section_prop['landscape']

        self.page_spec_name = self.section_prop['page-spec']
        self.page_spec = ConfigService()._page_specs['page-spec'][self.page_spec_name]

        self.margin_spec_name = self.section_prop['margin-spec']
        self.margin_spec = ConfigService()._page_specs['margin-spec'][self.margin_spec_name]

        # TODO: background-image is not used as of now
        # self.background_image = self.section_prop['background-image']

        self.bookmark = self.section_prop['bookmark']

        self.autocrop = self.section_prop['autocrop']
        self.page_bg = self.section_prop['page-bg']

        self.document_index = self.section_meta['document-index']
        self.document_name = self.section_meta['document-name']
        self.section_index = self.section_meta['section-index']
        self.section_name = self.section_meta['section-name']
        self.orientation = self.section_meta['orientation']
        self.first_section = self.section_meta['first-section']
        self.document_first_section = self.section_meta['document-first-section']
        self.different_firstpage = self.section_meta['different-firstpage']
        self.different_odd_even_pages = self.section_meta['different-odd-even-pages']
        self.document_nesting_depth = self.section_meta['document-nesting-depth']

        # TODO: page-layout is not used as of now
        # self.page_layout = self.section_meta['page-layout']

        self.section_id = f"D{str(self.document_index).zfill(3)}--S{str(self.section_index).zfill(3)}"

        self.master_page_name = f"mp-{self.section_id}"
        self.master_page = create_master_page(self._odt, first_section=self.first_section, document_index=self.document_index, master_page_name=self.master_page_name, page_spec=self.page_spec, margin_spec=self.margin_spec, orientation=self.orientation, nesting_level=nesting_level+1)
        self.master_page_name = self.master_page.getAttribute('name')

        # handle if first-page is different
        if self.different_firstpage:
            self.first_master_page_name = f"mp-{self.section_id}-first"
            self.first_master_page = create_master_page(self._odt, first_section=self.first_section, document_index=self.document_index, master_page_name=self.first_master_page_name, page_spec=self.page_spec, margin_spec=self.margin_spec, orientation=self.orientation, next_master_page_style=self.master_page_name, nesting_level=nesting_level+1)
            self.first_master_page_name = self.first_master_page.getAttribute('name')
        else:
            self.first_master_page_name = None
            self.first_master_page = None

        # handle page orientation
        if self.landscape:
            self.section_width = float(self.page_spec['height']) - float(self.margin_spec['left']) - float(self.margin_spec['right']) - float(self.margin_spec['gutter'])
            self.section_height = float(self.page_spec['width']) - float(self.margin_spec['top']) - float(self.margin_spec['bottom'])
        else:
            self.section_width = float(self.page_spec['width']) - float(self.margin_spec['left']) - float(self.margin_spec['right']) - float(self.margin_spec['gutter'])
            self.section_height = float(self.page_spec['height']) - float(self.margin_spec['top']) - float(self.margin_spec['bottom'])

        # handle headers and footers
        self.header_odd   = None
        self.footer_odd   = None
        self.header_even  = None
        self.footer_even  = None
        self.header_first = None
        self.footer_first = None

        if self._section_data['header-odd']:
            self.header_odd   = OdtPageHeaderFooter(self._section_data['header-odd'],   self.section_width, self.section_index, header_footer='header', odd_even='odd',   document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level)
    
        if self._section_data['footer-odd']:
            self.footer_odd   = OdtPageHeaderFooter(self._section_data['footer-odd'],   self.section_width, self.section_index, header_footer='footer', odd_even='odd',   document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level)

        if self._section_data['header-even']:
            self.header_even  = OdtPageHeaderFooter(self._section_data['header-even'],  self.section_width, self.section_index, header_footer='header', odd_even='even',  document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level)

        if self._section_data['footer-even']:
            self.footer_even  = OdtPageHeaderFooter(self._section_data['footer-even'],  self.section_width, self.section_index, header_footer='footer', odd_even='even',  document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level)

        if self._section_data['header-first']:
            self.header_first = OdtPageHeaderFooter(self._section_data['header-first'], self.section_width, self.section_index, header_footer='header', odd_even='first', document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level)

        if self._section_data['footer-first']:
            self.footer_first = OdtPageHeaderFooter(self._section_data['footer-first'], self.section_width, self.section_index, header_footer='footer', odd_even='first', document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level)

        self.section_contents = OdtContent(content_data=self._section_data.get('contents'), content_width=self.section_width, document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level)


    ''' get section heading
    '''
    def get_heading(self, nesting_level=0):
        style_attributes = {}

        # identify what style the heading will be and its content
        if not self.hide_heading:
            heading_text = self.heading
            if self.label != '':
                heading_text = f"{self.label} {heading_text}".strip()

            outline_level = self.level + self.document_nesting_depth
            if outline_level == 0:
                parent_style_name = 'Title'
            else:
                parent_style_name = f"Heading_20_{outline_level}"

        else:
            heading_text = ''
            parent_style_name = 'Text_20_body'
            outline_level = 0

        style_attributes['parentstylename'] = parent_style_name

        return heading_text, outline_level, style_attributes


    ''' Header/Footer processing
    '''
    def process_header_footer(self, nesting_level=0):
        if self.different_firstpage:
            master_page = self.first_master_page

            if self.header_first:
                self.header_first.page_header_footer_to_odt(odt=self._odt, master_page=master_page, nesting_level=nesting_level+1)

            if self.footer_first:
                self.footer_first.page_header_footer_to_odt(odt=self._odt, master_page=master_page, nesting_level=nesting_level+1)

        master_page = self.master_page

        if self.header_odd:
            self.header_odd.page_header_footer_to_odt(odt=self._odt, master_page=master_page, nesting_level=nesting_level+1)

        if self.footer_odd:
            self.footer_odd.page_header_footer_to_odt(odt=self._odt, master_page=master_page, nesting_level=nesting_level+1)

        if self.header_even:
            self.header_even.page_header_footer_to_odt(odt=self._odt, master_page=master_page, nesting_level=nesting_level+1)

        if self.footer_even:
            self.footer_even.page_header_footer_to_odt(odt=self._odt, master_page=master_page, nesting_level=nesting_level+1)


    ''' generates the odt code
    '''
    def section_to_odt(self, nesting_level=0):
        self.process_header_footer(nesting_level=nesting_level+1)

        #  get heading
        heading_text, outline_level, style_attributes = self.get_heading(nesting_level=nesting_level+1)

        paragraph_attributes = None
        # handle section-break and page-break
        if self.section_break:
            # if it is a new-section, we create a new paragraph-style based on parent_style_name with the master-page and apply it
            style_name = f"P{self.section_id}-P0-with-section-break"
            if self.different_firstpage:
                style_attributes['masterpagename'] = self.first_master_page_name
            else:
                style_attributes['masterpagename'] = self.master_page_name
        else:
            if self.page_break:
                # if it is a new-page, we create a new paragraph-style based on parent_style_name with the page-break and apply it
                paragraph_attributes = {'breakbefore': 'page'}
                style_name = f"P{self.section_id}-P0-with-page-break"
            else:
                style_name = f"P{self.section_id}-P0"

        style_attributes['name'] = style_name

        style_name = create_paragraph_style(self._odt, style_attributes=style_attributes, paragraph_attributes=paragraph_attributes)
        paragraph = create_paragraph(self._odt, style_name, text_content=heading_text, outline_level=outline_level, bookmark=self.bookmark)

        if self.heading_style:
            style = get_style_by_name(odt=self._odt, style_name=style_name)
            if style:
                if self.heading_style not in ConfigService()._style_specs:
                    warn(f"custom style [{self.heading_style}] not defined in style-specs", nesting_level=nesting_level)

                else:
                    trace(f"applying custom style [{self.heading_style}] to heading", nesting_level=nesting_level)
                    apply_custom_style(style=style, custom_properties=ConfigService()._style_specs[self.heading_style], nesting_level=nesting_level+1)

                    # handle background image
                    if 'page-background' in ConfigService()._style_specs[self.heading_style]:
                        for pb_dict in ConfigService()._style_specs[self.heading_style]['page-background']:
                            pb_image = InlineImage(ii_dict=pb_dict)
                            add_background_image_to_master_page(odt=self._odt, master_page=self.master_page, background_image_path=pb_image.file_path, nesting_level=nesting_level+1)

            else:
                warn(f"style [{style_name}] not found", nesting_level=nesting_level)

        self._odt.text.addElement(paragraph)



''' Odt table section object
'''
class OdtTableSection(OdtSectionBase):

    ''' constructor
    '''
    def __init__(self, odt, section_data, nesting_level=0):
        super().__init__(odt=odt, section_data=section_data, nesting_level=nesting_level)


    ''' generates the odt code
    '''
    def section_to_odt(self, nesting_level=0):
        super().section_to_odt(nesting_level=nesting_level+1)
        self.section_contents.content_to_odt(odt=self._odt, container=self._odt.text, nesting_level=nesting_level+1)



''' Odt gsheet section object
'''
class OdtGsheetSection(OdtSectionBase):

    ''' constructor
    '''
    def __init__(self, odt, section_data, nesting_level=0):
        super().__init__(odt=odt, section_data=section_data, nesting_level=nesting_level)


    ''' generates the odt code
    '''
    def section_to_odt(self, nesting_level=0):
        super().section_to_odt(nesting_level=nesting_level+1)

        # for embedded gsheets, 'contents' does not contain the actual content to render, rather we get a list of sections where each section contains the actual content
        if self._section_data['contents'] is not None and 'sections' in self._section_data['contents']:
            # process the sections
            section_list_to_odt(odt=self._odt, section_list=self._section_data['contents']['sections'], nesting_level=nesting_level+1)



''' Odt ToC section object
'''
class OdtToCSection(OdtSectionBase):

    ''' constructor
    '''
    def __init__(self, odt, section_data, nesting_level=0):
        super().__init__(odt=odt, section_data=section_data, nesting_level=nesting_level)


    ''' generates the odt code
    '''
    def section_to_odt(self, nesting_level=0):
        super().section_to_odt(nesting_level=nesting_level+1)
        toc = create_toc(nesting_level=nesting_level+1)
        if toc:
            self._odt.text.addElement(toc)



''' Odt LoT section object
'''
class OdtLoTSection(OdtSectionBase):

    ''' constructor
    '''
    def __init__(self, odt, section_data, nesting_level=0):
        super().__init__(odt=odt, section_data=section_data, nesting_level=nesting_level)


    ''' generates the odt code
    '''
    def section_to_odt(self, nesting_level=0):
        super().section_to_odt(nesting_level=nesting_level+1)
        toc = create_lot(nesting_level=nesting_level+1)
        if toc:
            self._odt.text.addElement(toc)



''' Odt LoF section object
'''
class OdtLoFSection(OdtSectionBase):

    ''' constructor
    '''
    def __init__(self, odt, section_data, nesting_level=0):
        super().__init__(odt=odt, section_data=section_data, nesting_level=nesting_level)


    ''' generates the odt code
    '''
    def section_to_odt(self, nesting_level=0):
        super().section_to_odt(nesting_level=nesting_level+1)
        toc = create_lof(nesting_level=nesting_level+1)
        if toc:
            self._odt.text.addElement(toc)



''' Odt Pdf section object
'''
class OdtPdfSection(OdtSectionBase):

    ''' constructor
    '''
    def __init__(self, odt, section_data, nesting_level=0):
        super().__init__(odt=odt, section_data=section_data, nesting_level=nesting_level)


    ''' generates the odt code
    '''
    def section_to_odt(self, nesting_level=0):
        super().section_to_odt(nesting_level=nesting_level+1)

        # the images go one after another
        text_attributes = {'fontsize': 2}
        style_attributes = {}
        if 'contents' in self._section_data:
            if self._section_data['contents'] and 'images' in self._section_data['contents']:
                for i, image in enumerate(self._section_data['contents']['images']):
                    # we need to set bookmark for each pdf page if the section has a bookmark arrached with it, we just append page number
                    this_image_bookmark = None
                    if self.bookmark:
                        key = list(self.bookmark)[0]
                        this_image_bookmark = f"{key}.{str(i).zfill(3)}"

                    paragraph_attributes = {'textalign': TEXT_HALIGN_MAP['CENTER']}
                    if True:
                        paragraph_attributes['breakbefore'] = 'page'

                    # if the image should be treated as page bg
                    if self.page_bg:
                        master_page_name = f"{self.master_page_name}-page-{str(i).zfill(3)}"
                        self.master_page = create_master_page(self._odt, first_section=self.first_section, document_index=self.document_index, master_page_name=master_page_name, page_spec=self.page_spec, margin_spec=self.margin_spec, orientation=self.orientation, nesting_level=nesting_level+1)
                        master_page_name = self.master_page.getAttribute('name')
                        # handle background image
                        add_background_image_to_master_page(odt=self._odt, master_page=self.master_page, background_image_path=image['path'], nesting_level=nesting_level+1)

                        paragraph_style_name = f"{self.master_page_name}-P-{str(i).zfill(3)}"
                        paragraph = create_paragraph_with_masterpage(odt=self._odt, style_name=paragraph_style_name, master_page_name=master_page_name, nesting_level=nesting_level+1)
                        if this_image_bookmark:
                            paragraph.addElement(text.Bookmark(name=this_image_bookmark))
                        
                        self._odt.text.addElement(paragraph)

                        self.process_header_footer()

                    # if the image should be inserted in page body area
                    else:
                        # we need to create an inline-image object
                        ii_dict = {
                            'file-path': image['path'],
                            'image-width': image['width'],
                            'image-height': image['height'],
                            'type': 'inline',
                            'fit-height-to-container': True,
                            # there is a quirk here valign is not middle, it is center
                            'position': 'center center',
                        }

                        inline_image = InlineImage(ii_dict=ii_dict)

                        graphic_properties_attributes = inline_image.graphic_properties_attributes(nesting_level=nesting_level+1)
                        frame_attributes = inline_image.frame_attributes(container_width=self.section_width, container_height=self.section_height - PDF_PAGE_HEIGHT_OFFSET, nesting_level=nesting_level+1)
                        draw_frame = create_image_frame(odt=self._odt, picture_path=image['path'], frame_attributes=frame_attributes, graphic_properties_attributes=graphic_properties_attributes, nesting_level=nesting_level+1)

                        style_name = create_paragraph_style(self._odt, style_attributes=style_attributes, paragraph_attributes=paragraph_attributes, text_attributes=text_attributes, nesting_level=nesting_level+1)
                        paragraph = create_paragraph(self._odt, style_name, nesting_level=nesting_level+1)
                        paragraph.addElement(draw_frame)

                        if this_image_bookmark:
                            paragraph.addElement(text.Bookmark(name=this_image_bookmark))

                        self._odt.text.addElement(paragraph)



''' Odt section content base object
'''
class OdtContent(object):

    ''' constructor
    '''
    def __init__(self, content_data, content_width, document_nesting_depth, nesting_level=0):
        self.content_data = content_data
        self.content_width = content_width
        self.document_nesting_depth = document_nesting_depth

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
                            new_row = Row(row_num=r, row_data=row_data, section_width=self.content_width, column_widths=self.column_widths, row_height=self.row_metadata_list[r].inches, document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level+1)
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
    def process(self, nesting_level=0):

        # first we identify the missing cells or blank cells for merged spans
        for merge in self.merge_list:
            first_row = merge.start_row
            first_col = merge.start_col
            last_row = merge.end_row
            last_col = merge.end_col
            row_span = merge.row_span
            col_span = merge.col_span
            first_row_object = self.cell_matrix[first_row]
            first_cell = first_row_object.get_cell(first_col, nesting_level=nesting_level+1)

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
                        next_cell_in_row = Cell(row_num=r+2, col_num=c, value=None, column_widths=first_cell.column_widths, row_height=row_height, document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level+1)
                        next_row_object.insert_cell(c, next_cell_in_row, nesting_level=nesting_level+1)

                    if next_cell_in_row.is_empty:
                        # the cell is a newly inserted one, its format should be the same (for borders, colors) as the first cell so that we can draw borders properly
                        next_cell_in_row.copy_format_from(first_cell, nesting_level=nesting_level+1)

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
    def split(self, nesting_level=0):
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
        container may be odt.text or a table-cell
    '''
    def content_to_odt(self, odt, container, nesting_level=0):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        # iterate through tables and blocks contents
        for block in self.content_list:
            block.block_to_odt(odt=odt, container=container, nesting_level=nesting_level)



''' Odt Page Header Footer object
'''
class OdtPageHeaderFooter(OdtContent):

    ''' constructor
        header_footer : header/footer
        odd_even      : first/odd/even(left)
    '''
    def __init__(self, content_data, section_width, section_index, header_footer, odd_even, document_nesting_depth, nesting_level=0):
        super().__init__(content_data, section_width, document_nesting_depth=document_nesting_depth, nesting_level=nesting_level)
        self.header_footer, self.odd_even = header_footer, odd_even
        self.id = f"{self.header_footer}-{self.odd_even}-{section_index}"


    ''' generates the odt code
    '''
    def page_header_footer_to_odt(self, odt, master_page, nesting_level=0):
        if self.content_data is None:
            return

        header_footer = create_header_footer(odt=odt, master_page=master_page, header_or_footer=self.header_footer, odd_or_even=self.odd_even, nesting_level=nesting_level+1)
        if header_footer:
            # iterate through tables and blocks contents
            for block in self.content_list:
                block.block_to_odt(odt=odt, container=header_footer, nesting_level=nesting_level+1)



''' Odt Block object wrapper base class (plain odt, table, header etc.)
'''
class OdtBlock(object):

    ''' constructor
    '''
    def __init__(self, nesting_level=0):
        pass



''' Odt Table object wrapper
'''
class OdtTable(OdtBlock):

    ''' constructor
    '''
    def __init__(self, cell_matrix, start_row, end_row, column_widths, nesting_level=0):
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


    ''' generates the odt code
    '''
    def block_to_odt(self, odt, container, nesting_level=0):
        # debug(f". {self.__class__.__name__} : {inspect.stack()[0][3]}")

        # create the table with styles
        table_style_attributes = {'name': f"{self.table_name}_style"}
        table_properties_attributes = {'width': f"{sum(self.column_widths)}in"}
        table = create_table(odt, self.table_name, table_style_attributes=table_style_attributes, table_properties_attributes=table_properties_attributes, nesting_level=nesting_level+1)

        # table-columns
        for c in range(0, len(self.column_widths)):
            col_a1 = COLUMNS[c]
            col_width = self.column_widths[c]
            table_column_name = f"{self.table_name}.{col_a1}"
            table_column_style_attributes = {'name': f"{table_column_name}_style"}
            table_column_properties_attributes = {'columnwidth': f"{col_width}in", 'useoptimalcolumnwidth': False}
            table_column = create_table_column(odt, table_column_name, table_column_style_attributes, table_column_properties_attributes, nesting_level=nesting_level+1)
            table.addElement(table_column)

        # iterate header rows render the table's contents
        table_header_rows = create_table_header_rows(nesting_level=nesting_level+1)
        for row in self.table_cell_matrix[0:self.header_row_count]:
            table_row = row.row_to_odt_table_row(odt, self.table_name)
            table_header_rows.addElement(table_row)

        table.addElement(table_header_rows)

        # iterate rows and cells to render the table's contents
        num_cols = len(self.column_widths)
        num_rows = len(self.table_cell_matrix)
        for row in self.table_cell_matrix[self.header_row_count:]:
            table_row = row.row_to_odt_table_row(odt, self.table_name)
            table.addElement(table_row)

        container.addElement(table)



''' Odt Block object wrapper
'''
class OdtParagraph(OdtBlock):

    ''' constructor
    '''
    def __init__(self, data_row, row_number, nesting_level=0):
        self.data_row = data_row
        self.row_number = row_number

    ''' string representation
    '''
    def __repr__(self):
        return f"__para__"

    """ generates the odt code
    """
    def block_to_odt(self, odt, container, nesting_level=0):
        # generate the block, only the first cell of the data_row to be produced
        if len(self.data_row.cells) > 0:
            # We take the first cell, the cell will take the whole row width
            cell_to_produce = self.data_row.get_cell(0)
            cell_to_produce.cell_width = sum(cell_to_produce.column_widths)
            cell_to_produce.cell_to_odt(odt=odt, container=container, nesting_level=nesting_level+1)



#   ----------------------------------------------------------------------------------------------------------------
#   gsheet row and cell wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' gsheet Row object wrapper
'''
class Row(object):

    ''' constructor
    '''
    def __init__(self, row_num, row_data, section_width, column_widths, row_height, document_nesting_depth, nesting_level=0):
        self.row_num, self.section_width, self.column_widths, self.row_height, self.document_nesting_depth = row_num, section_width, column_widths, row_height, document_nesting_depth
        self.row_id = f"{self.row_num+1}"
        self.row_name = f"row: [{self.row_id}]"

        self.cells = []
        c = 0
        values = row_data.get('values', [])
        if len(values) > 1:
            for value in values[1:]:
                self.cells.append(Cell(row_num=self.row_num, col_num=c, value=value, column_widths=self.column_widths, row_height=self.row_height, document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level+1))
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
    def preprocess_row(self, nesting_level=0):
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
    def is_empty(self, nesting_level=0):
        return (len(self.cells) == 0)


    ''' gets a specific cell by ordinal
    '''
    def get_cell(self, c, nesting_level=0):
        if c >= 0 and c < len(self.cells):
            return self.cells[c]
        else:
            return None


    ''' inserts a specific cell at a specific ordinal
    '''
    def insert_cell(self, pos, cell, nesting_level=0):
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
        1. it contains a note {'content': 'out-of-cell/free'}
        2. it contains a note {'style': '...'} and it is the only non-empty cell in the row
    '''
    def is_free_content(self, nesting_level=0):
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
    def is_table_start(self, nesting_level=0):
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


    ''' generates the odt code
    '''
    def row_to_odt_table_row(self, odt, table_name, nesting_level=0):
        self.table_name = table_name

        # create table-row
        table_row_style_attributes = {'name': f"{self.table_name}-{self.row_num}"}
        row_height = f"{self.row_height}in"
        table_row_properties_attributes = {'keeptogether': True, 'minrowheight': row_height, 'useoptimalrowheight': True}
        table_row = create_table_row(odt, table_row_style_attributes, table_row_properties_attributes, nesting_level=nesting_level+1)

        # iterate over the cells
        c = 0
        for cell in self.cells:
            if cell is None:
                warn(f"{self.row_name} has a Null cell at {c}", nesting_level=nesting_level)
            else:
                table_cell = cell.cell_to_odt_table_cell(odt, self.table_name, nesting_level=nesting_level+1)
                if table_cell:
                    table_row.addElement(table_cell)

                else:
                    warn(f"Invalid table-cell : [{cell.cell_id}]", nesting_level=nesting_level)

            c = c + 1

        return table_row



''' gsheet Cell object wrapper
'''
class Cell(object):

    ''' constructor
    '''
    def __init__(self, row_num, col_num, value, column_widths, row_height, document_nesting_depth, nesting_level=0):
        self.row_num, self.col_num, self.column_widths, self.document_nesting_depth  = row_num, col_num, column_widths, document_nesting_depth
        self.cell_id = f"{COLUMNS[self.col_num+1]}{self.row_num+1}"
        self.cell_name = self.cell_id
        self.value = value
        self.text_format_runs = []
        self.cell_width = self.column_widths[self.col_num]
        self.cell_height = row_height

        self.merge_spec = CellMergeSpec()
        self.note = None
        self.effective_format = CellFormat(format_dict=None)
        self.user_entered_format = CellFormat(format_dict=None)

        # considering merges, we have effective cell width and height
        self.effective_cell_width = self.cell_width
        self.effective_cell_height = self.cell_height

        # a cell may have multiple images as inline images
        self.inline_images = []
        self.image_frames = []

        if self.value:
            if 'inline-image' in self.value:
                for ii_dict in self.value.get('inline-image', []):
                    self.inline_images.append(InlineImage(ii_dict, nesting_level=nesting_level+1))
                
            note_dict = self.value.get('notes', {})
            self.note = CellNote(document_nesting_depth=self.document_nesting_depth, note_dict=note_dict, nesting_level=nesting_level+1)

            self.formatted_value = self.value.get('formattedValue', '')

            self.effective_format = CellFormat(self.value.get('effectiveFormat'), nesting_level=nesting_level+1)

            for text_format_run in self.value.get('textFormatRuns', []):
                self.text_format_runs.append(TextFormatRun(run_dict=text_format_run, default_format=self.effective_format.text_format.source, nesting_level=nesting_level+1))

            # presence of userEnteredFormat makes the cell non-empty
            if 'userEnteredFormat' in self.value:
                self.user_entered_format = CellFormat(format_dict=self.value.get('userEnteredFormat'), nesting_level=nesting_level+1)
                self.is_empty = False

                # HACK: handle background-color - if user_entered_format does not have backgroundColor, omit backgroundColor from effective_format
                if 'backgroundColor' in self.value.get('userEnteredFormat'):
                    self.effective_format.bgcolor = self.user_entered_format.bgcolor

            else:
                self.user_entered_format = None
                self.is_empty = True

            # we need to identify exactly what kind of value the cell contains
            if 'contents' in self.value:
                self.cell_value = ContentValue(effective_format=self.effective_format, content_value=self.value['contents'], document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level+1)

            elif 'userEnteredValue' in self.value:
                if 'image' in self.value['userEnteredValue']:
                    self.cell_value = ImageValue(effective_format=self.effective_format, image_value=self.value['userEnteredValue']['image'], document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level+1)

                else:
                    if len(self.text_format_runs):
                        self.cell_value = TextRunValue(effective_format=self.effective_format, text_format_runs=self.text_format_runs, formatted_value=self.formatted_value, document_nesting_depth=self.document_nesting_depth, outline_level=self.note.outline_level, keep_line_breaks=self.note.keep_line_breaks, nesting_level=nesting_level+1)

                    else:
                        if self.note.script and self.note.script == 'latex':
                            self.cell_value = LatexValue(effective_format=self.effective_format, string_value=self.value['userEnteredValue'], formatted_value=self.formatted_value, document_nesting_depth=self.document_nesting_depth, outline_level=self.note.outline_level, nesting_level=nesting_level+1)

                        else:
                            self.cell_value = StringValue(effective_format=self.effective_format, string_value=self.value['userEnteredValue'], formatted_value=self.formatted_value, document_nesting_depth=self.document_nesting_depth, outline_level=self.note.outline_level, bookmark=self.note.bookmark, keep_line_breaks=self.note.keep_line_breaks, nesting_level=nesting_level+1)

            else:
                self.cell_value = StringValue(effective_format=self.effective_format, string_value='', formatted_value=self.formatted_value, document_nesting_depth=self.document_nesting_depth, outline_level=self.note.outline_level, bookmark=self.note.bookmark, keep_line_breaks=self.note.keep_line_breaks, nesting_level=nesting_level+1)

        else:
            # value can have a special case it can be an empty ditionary when the cell is an inner cell of a column merge
            self.cell_value = None
            self.formatted_value = ''
            self.is_empty = True


    ''' string representation
    '''
    def __repr__(self):
        s = f"{self.cell_name:>4}, value: {not self.is_empty:<1}, mr: {self.merge_spec.multi_row:<9}, mc: {self.merge_spec.multi_col:<9} [{self.formatted_value[0:50]}]"
        # s = f"{self.cell_name:>4}, value: {not self.is_empty:<1}, mr: {self.merge_spec.multi_row:<9}, mc: {self.merge_spec.multi_col:<9} [{self.effective_format.borders}]"
        return s


    ''' odt code for cell content
    '''
    def cell_to_odt_table_cell(self, odt, table_name, nesting_level=0):
        self.table_name = table_name
        col_a1 = COLUMNS[self.col_num+1]
        table_cell_style_attributes = {'name': f"{self.table_name}.{col_a1}{self.row_num+1}_style"}

        table_cell_properties_attributes = {}
        if self.effective_format:
            table_cell_properties_attributes = self.effective_format.table_cell_attributes(self.merge_spec, nesting_level=nesting_level+1)
        else:
            table_cell_properties_attributes = {}
            warn(f"{self} : NO effective_format")

        if not self.is_covered():
            # wrap this into a table-cell
            table_cell_attributes = self.merge_spec.table_cell_attributes(nesting_level=nesting_level+1)

            # handle cell background image
            background_image_style = None

            # the cell may have a custom style with a bg image 
            if self.note.style is not None and self.note.style in ConfigService()._style_specs:
                # this custom style may have an inline-image, if so apply it
                if 'inline-image' in ConfigService()._style_specs[self.note.style]:
                    for ii_dict in ConfigService()._style_specs[self.note.style]['inline-image']:
                        inline_image = InlineImage(ii_dict, nesting_level=nesting_level+1)

                        # consider only the first bg image
                        if inline_image.type == 'background':
                            background_image_style = create_background_image_style(odt=odt, picture_path=inline_image.file_path, nesting_level=nesting_level+1)
                            break

                        else:
                            self.inline_images.append(inline_image)

            # and/or there might me inline images in notes
            for inline_image in self.inline_images:
                # now if the background does not have any position, it is to be trated as a background image
                if inline_image.type == 'background':
                    background_image_style = create_background_image_style(odt=odt, picture_path=inline_image.file_path, nesting_level=nesting_level+1)

                # the image is positioned, it is to be positioned as a non-bg image
                elif inline_image.type == 'inline':
                    graphic_properties_attributes = inline_image.graphic_properties_attributes(nesting_level=nesting_level+1)
                    frame_attributes = inline_image.frame_attributes(container_width=self.effective_cell_width, container_height=self.effective_cell_height, nesting_level=nesting_level+1)
                    self.image_frames.append(create_image_frame(odt=odt, picture_path=inline_image.file_path, frame_attributes=frame_attributes, graphic_properties_attributes=graphic_properties_attributes, nesting_level=nesting_level+1))

                else:
                    warn(f"invalid inline-image type [{inline_image.type}]", nesting_level=nesting_level)

            table_cell = create_table_cell(odt, table_cell_style_attributes, table_cell_properties_attributes, table_cell_attributes, background_image_style=background_image_style, nesting_level=nesting_level+1)

            if table_cell:
                self.cell_to_odt(odt=odt, container=table_cell, is_table_cell=True, nesting_level=nesting_level+1)

        else:
            # wrap this into a covered-table-cell
            table_cell = create_covered_table_cell(odt, table_cell_style_attributes, table_cell_properties_attributes, nesting_level=nesting_level+1)

        return table_cell


    ''' odt code for cell content
    '''
    def cell_to_odt(self, odt, container, is_table_cell=False, nesting_level=0):
        # trace(f"{self}", nesting_level=nesting_level)
        if self.note:
            paragraph_attributes_from_notes = self.note.paragraph_attributes(nesting_level=nesting_level+1)
            style_attributes = self.note.style_attributes(nesting_level=nesting_level+1)
            footnote_list = self.note.footnotes
            force_halign = self.note.force_halign
            angle = self.note.angle
        else:
            paragraph_attributes_from_notes = {}
            style_attributes = {}
            footnote_list = {}
            force_halign = False
            angle = 0

        if self.effective_format:
            paragraph_attributes_from_effective_format = self.effective_format.paragraph_attributes(is_table_cell=is_table_cell, cell_merge_spec=self.merge_spec, force_halign=force_halign, nesting_level=nesting_level+1)
            text_attributes = self.effective_format.text_attributes(angle, nesting_level=nesting_level+1)
        else:
            paragraph_attributes_from_effective_format = {}
            text_attributes = {}

        paragraph_attributes = {**paragraph_attributes_from_notes,  **paragraph_attributes_from_effective_format}


        # for string and image it returns a paragraph, for embedded content a list
        # the content is not valid for multirow LastCell and InnerCell
        if self.merge_spec.multi_row in [MultiSpan.No, MultiSpan.FirstCell] and self.merge_spec.multi_col in [MultiSpan.No, MultiSpan.FirstCell]:
            if self.cell_value:
                paragraph = self.cell_value.value_to_odt(odt, container=container, container_width=self.effective_cell_width, container_height=self.effective_cell_height, style_attributes=style_attributes, paragraph_attributes=paragraph_attributes, text_attributes=text_attributes, footnote_list=footnote_list, bookmark=self.note.bookmark, nesting_level=nesting_level+1)
                # place the image frame
                if paragraph:
                    for image_frame in self.image_frames:
                        paragraph.addElement(image_frame)

                    # the cell/paragraph may have custom style, if so apply it
                    if self.note.style is not None and self.note.style in ConfigService()._style_specs:
                        trace(f"applying custom style [{self.note.style}] to {self}", nesting_level=nesting_level)
                        style = get_style_by_name(odt=odt, style_name=paragraph.getAttribute("stylename"), nesting_level=nesting_level+1)
                        apply_custom_style(style=style, custom_properties=ConfigService()._style_specs[self.note.style], nesting_level=nesting_level+1)

                else:
                    # warn(f"[{self}] No paragraph to add inline image(s) to", nesting_level=nesting_level)
                    pass


    ''' Copy format from the cell passed
    '''
    def copy_format_from(self, from_cell, nesting_level=0):
        self.effective_format = from_cell.effective_format


    ''' is the cell part of a merge
    '''
    def is_covered(self, nesting_level=0):
        if self.merge_spec.multi_col in [MultiSpan.InnerCell, MultiSpan.LastCell] or self.merge_spec.multi_row in [MultiSpan.InnerCell, MultiSpan.LastCell]:
            return True

        else:
            return False



''' gsheet cell value object wrapper
'''
class CellValue(object):

    ''' constructor
    '''
    def __init__(self, effective_format, document_nesting_depth, outline_level=0):
        self.effective_format = effective_format
        self.document_nesting_depth = document_nesting_depth
        self.outline_level = outline_level



''' string type CellValue
'''
class StringValue(CellValue):

    ''' constructor
    '''
    def __init__(self, effective_format, string_value, formatted_value, document_nesting_depth, outline_level=0, bookmark={}, keep_line_breaks=False, directives=True, nesting_level=0):
        super().__init__(effective_format=effective_format, document_nesting_depth=document_nesting_depth, outline_level=outline_level)
        if formatted_value:
            self.value = formatted_value
        else:
            if string_value and 'stringValue' in string_value:
                self.value = string_value['stringValue']
            else:
                self.value = ''

        self.keep_line_breaks = keep_line_breaks
        self.bookmark = bookmark
        self.directives = directives


    ''' string representation
    '''
    def __repr__(self):
        s = f"string : [{self.value}]"
        return s


    ''' generates the odt code
    '''
    def value_to_odt(self, odt, container, container_width, container_height, style_attributes, paragraph_attributes, text_attributes, footnote_list, bookmark, nesting_level=0):
        if container is None:
            container = odt.text

        style_name = create_paragraph_style(odt, style_attributes=style_attributes, paragraph_attributes=paragraph_attributes, text_attributes=text_attributes, nesting_level=nesting_level+1)
        paragraph = create_paragraph(odt, style_name, text_content=self.value, outline_level=self.outline_level, footnote_list=footnote_list, bookmark=bookmark, keep_line_breaks=self.keep_line_breaks, directives=self.directives, nesting_level=nesting_level+1)
        container.addElement(paragraph)

        return paragraph



''' LaTex type CellValue
'''
class LatexValue(CellValue):

    ''' constructor
    '''
    def __init__(self, effective_format, string_value, formatted_value, document_nesting_depth, outline_level=0, nesting_level=0):
        super().__init__(effective_format=effective_format, document_nesting_depth=document_nesting_depth, outline_level=outline_level)
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


    ''' generates the odt code
    '''
    def value_to_odt(self, odt, container, container_width, container_height, style_attributes, paragraph_attributes, text_attributes, footnote_list, bookmark, nesting_level=0):
        if container is None:
            container = odt.text

        style_name = create_paragraph_style(odt, style_attributes=style_attributes, paragraph_attributes=paragraph_attributes, text_attributes=text_attributes, nesting_level=nesting_level+1)
        paragraph = create_mathml(odt, style_name, latex_content=self.value, nesting_level=nesting_level+1)
        container.addElement(paragraph)

        return paragraph



''' text-run type CellValue
'''
class TextRunValue(CellValue):

    ''' constructor
    '''
    def __init__(self, effective_format, text_format_runs, formatted_value, document_nesting_depth, outline_level=0, keep_line_breaks=False, nesting_level=0):
        super().__init__(effective_format=effective_format, document_nesting_depth=document_nesting_depth, outline_level=outline_level)
        self.text_format_runs = text_format_runs
        self.formatted_value = formatted_value
        self.keep_line_breaks = keep_line_breaks

    ''' string representation
    '''
    def __repr__(self):
        s = f"text-run : [{self.formatted_value}]"
        return s

    ''' generates the odt code
    '''
    def value_to_odt(self, odt, container, container_width, container_height, style_attributes, paragraph_attributes, text_attributes, footnote_list, bookmark, nesting_level=0):
        if container is None:
            container = odt.text

        run_value_list = []
        processed_idx = len(self.formatted_value)
        for text_format_run in reversed(self.text_format_runs):
            text = self.formatted_value[:processed_idx]
            run_value_list.insert(0, text_format_run.text_attributes(text))
            processed_idx = text_format_run.start_index

        style_name = create_paragraph_style(
            odt,
            style_attributes=style_attributes,
            paragraph_attributes=paragraph_attributes,
            text_attributes=text_attributes,
            nesting_level=nesting_level+1
        )
        paragraph = create_paragraph(odt, style_name, run_list=run_value_list, outline_level=self.outline_level, footnote_list=footnote_list, bookmark=bookmark, keep_line_breaks=self.keep_line_breaks, nesting_level=nesting_level+1)
        container.addElement(paragraph)

        return paragraph



''' image type CellValue
'''
class ImageValue(CellValue):

    ''' constructor
    '''
    def __init__(self, effective_format, image_value, document_nesting_depth, outline_level=0, nesting_level=0):
        super().__init__(effective_format=effective_format, document_nesting_depth=document_nesting_depth, outline_level=outline_level)
        self.value = image_value


    ''' string representation
    '''
    def __repr__(self):
        s = f"image : [{self.value}]"
        return s


    ''' generates the odt code
    '''
    def value_to_odt(self, odt, container, container_width, container_height, style_attributes, paragraph_attributes, text_attributes, footnote_list, bookmark, nesting_level=0):
        if container is None:
            container = odt.text

        # even now the width may exceed actual cell width, we need to adjust for that
        dpi_x = DPI if self.value['dpi'][0] == 0 else self.value['dpi'][0]
        dpi_y = DPI if self.value['dpi'][1] == 0 else self.value['dpi'][1]
        image_width_in_pixel = self.value['size'][0]
        image_height_in_pixel = self.value['size'][1]
        image_width_in_inches =  image_width_in_pixel / dpi_x
        image_height_in_inches = image_height_in_pixel / dpi_y

        container_ratio = container_width/container_height
        image_ratio = image_width_in_inches/image_height_in_inches

        fit_height_to_container = False
        fit_width_to_container = False
        # mode 1 resizes the image to fit inside the cell, maintaining aspect ratio.
        # mode 2 stretches or compresses the image to fit inside the cell, ignoring aspect ratio.
        # mode 3 leaves the image at original size, which may cause cropping.
        # mode 4 allows the specification of a custom size
        if self.value['mode'] in [1, 2, 4]:
            # image is to be scaled within the cell width and height
            if container_ratio > image_ratio:
                fit_height_to_container = True
            else:
                fit_width_to_container = True

        else:
            pass

        text_attributes['fontsize'] = 2
        style_name = create_paragraph_style(odt, style_attributes=style_attributes, paragraph_attributes=paragraph_attributes, text_attributes=text_attributes, nesting_level=nesting_level+1)
        paragraph = create_paragraph(odt, style_name, bookmark=bookmark, nesting_level=nesting_level+1)

        # we need to create an inline-image object
        picture_path = self.value['path']
        ii_dict = {
            'file-path': picture_path,
            'image-width': image_width_in_inches,
            'image-height': image_height_in_inches,
            'type': 'inline',
            'fit-height-to-container': fit_height_to_container,
            'fit-width-to-container': fit_width_to_container,
            'keep-aspect-ratio': True,
            'position': f"{IMAGE_POSITION_HORIZONRAL[self.effective_format.halign.halign]} {IMAGE_POSITION_VERTICAL[self.effective_format.valign.valign]}",
            'wrap': 'run-through'
        }

        inline_image = InlineImage(ii_dict=ii_dict, nesting_level=nesting_level+1)
        graphic_properties_attributes = inline_image.graphic_properties_attributes(nesting_level=nesting_level+1)
        frame_attributes = inline_image.frame_attributes(container_width=container_width, container_height=container_height, nesting_level=nesting_level+1)
        # frame_attributes = inline_image.frame_attributes(nesting_level=nesting_level+1)
        image_frame = create_image_frame(odt=odt, picture_path=inline_image.file_path, frame_attributes=frame_attributes, graphic_properties_attributes=graphic_properties_attributes, nesting_level=nesting_level+1)

        paragraph.addElement(image_frame)
        container.addElement(paragraph)

        return paragraph



''' content type CellValue
'''
class ContentValue(CellValue):

    ''' constructor
    '''
    def __init__(self, effective_format, content_value, document_nesting_depth, outline_level=0, nesting_level=0):
        super().__init__(effective_format=effective_format, document_nesting_depth=document_nesting_depth, outline_level=outline_level)
        self.value = content_value


    ''' string representation
    '''
    def __repr__(self):
        s = f"content : [{self.value['sheets'][0]['properties']['title']}]"
        return s


    ''' generates the odt code
    '''
    def value_to_odt(self, odt, container, container_width, container_height, style_attributes, paragraph_attributes, text_attributes, footnote_list, bookmark, nesting_level=0):
        self.contents = OdtContent(content_data=self.value, content_width=container_width, document_nesting_depth=self.document_nesting_depth, nesting_level=nesting_level+1)
        self.contents.content_to_odt(odt=odt, container=container, nesting_level=nesting_level+1)

        return None



#   ----------------------------------------------------------------------------------------------------------------
#   gsheet cell property wrappers
#   ----------------------------------------------------------------------------------------------------------------

''' gsheet text format object wrapper
'''
class TextFormat(object):

    ''' constructor
    '''
    def __init__(self, text_format_dict=None, nesting_level=0):
        self.source = text_format_dict
        if self.source:
            self.fgcolor = RgbColor(text_format_dict.get('foregroundColor'))
            if 'fontFamily' in text_format_dict:
                self.font_family = text_format_dict['fontFamily']

            # else:
            #     self.font_family = ''


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
    def text_attributes(self, nesting_level=0):
        attributes = {}

        attributes['color'] = self.fgcolor.value()
        if self.font_family:
            attributes['fontname'] = self.font_family
            attributes['fontnameasian'] = self.font_family
            # attributes['fontnamecomplex'] = self.font_family

        attributes['fontsize'] = self.font_size
        attributes['fontsizeasian'] = self.font_size
        attributes['fontsizecomplex'] = self.font_size

        if self.is_bold:
            attributes['fontweight'] = "bold"
            attributes['fontweightasian'] = "bold"
            attributes['fontweightcomplex'] = "bold"

        if self.is_italic:
            attributes['fontstyle'] = "italic"
            attributes['fontstyleasian'] = "italic"
            attributes['fontstylecomplex'] = "italic"

        if self.is_underline:
            attributes['textunderlinestyle'] = "solid"
            attributes['textunderlinewidth'] = "auto"
            attributes['textunderlinecolor'] = "font-color"

        if self.is_strikethrough:
            attributes['textlinethroughstyle'] = "solid"
            attributes['textlinethroughtype'] = "single"

        return attributes



''' gsheet cell format object wrapper
'''
class CellFormat(object):

    ''' constructor
    '''
    def __init__(self, format_dict, nesting_level=0):
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
            self.text_rotation = TextRotation(format_dict.get("textRotation"))

            bgcolor_style_dict = format_dict.get('backgroundColorStyle')
            if bgcolor_style_dict:
                self.bgcolor_style = RgbColor(bgcolor_style_dict.get('rgbColor'))


    ''' attributes dict for Cell Text
    '''
    def text_attributes(self, angle=0, nesting_level=0):
        attributes = {}
        if self.text_format:
            attributes = self.text_format.text_attributes()

        if self.text_rotation and self.text_rotation.angle != 0:
            attributes["textrotationangle"] = self.text_rotation.angle

        if angle != 0:
            attributes['textrotationangle'] = angle

        return attributes


    ''' attributes dict for TableCellProperties
    '''
    def table_cell_attributes(self, cell_merge_spec, nesting_level=0):
        attributes = {}

        if self.valign:
            attributes['verticalalign'] = self.valign.valign

        if self.bgcolor:
            if self.bgcolor_style:
                if not self.bgcolor_style.is_white():
                    attributes['backgroundcolor'] = self.bgcolor_style.value()
            else:
                attributes['backgroundcolor'] = self.bgcolor.value()

        if self.wrapping:
            attributes['wrapoption'] = self.wrapping.wrapping

        borders_attributes = {}
        padding_attributes = {}
        if self.borders:
            borders_attributes = self.borders.table_cell_attributes(cell_merge_spec)

        if self.padding:
            padding_attributes = self.padding.table_cell_attributes()

        return {**attributes, **borders_attributes, **padding_attributes}


    ''' attributes dict for ParagraphProperties
    '''
    def paragraph_attributes(self, is_table_cell, cell_merge_spec, force_halign, nesting_level=0):
        # if the is left aligned, we do not set attribute to let the parent style determine what the alignment should be
        if force_halign:
            attributes = {'textalign': self.halign.halign}

        else:
            if self.halign is None or self.halign.halign in ['left']:
                attributes = {}
            else:
                attributes = {'textalign': self.halign.halign}

        if self.bgcolor:
            if self.bgcolor_style:
                if not self.bgcolor_style.is_white():
                    attributes['backgroundcolor'] = self.bgcolor_style.value()
            else:
                attributes['backgroundcolor'] = self.bgcolor.value()

        if self.valign:
            attributes['verticalalign'] = self.valign.valign

        borders_attributes = {}
        padding_attributes = {}
        if is_table_cell:
            pass

        else:
            # TODO: borders for out-of-cell-paragraphs
            if self.borders:
                borders_attributes = self.borders.paragraph_attributes()

            if self.padding:
                padding_attributes = self.padding.table_cell_attributes()

        return {**attributes, **borders_attributes, **padding_attributes}



''' gsheet cell borders object wrapper
'''
class Borders(object):

    ''' constructor
    '''
    def __init__(self, borders_dict=None, nesting_level=0):
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
    def table_cell_attributes(self, cell_merge_spec, nesting_level=0):
        attributes = {}

        # top and bottom
        if cell_merge_spec.multi_row in [MultiSpan.No, MultiSpan.FirstCell]:
            if self.top:
                attributes['bordertop'] = self.top.value()

            if self.bottom:
                attributes['borderbottom'] = self.bottom.value()

        if cell_merge_spec.multi_row in [MultiSpan.LastCell]:
            if self.bottom:
                attributes['borderbottom'] = self.bottom.value()


        # left and right
        if cell_merge_spec.multi_col in [MultiSpan.No, MultiSpan.FirstCell]:
            if self.left:
                attributes['borderleft'] = self.left.value()

            if self.right:
                attributes['borderright'] = self.right.value()

        if cell_merge_spec.multi_col in [MultiSpan.LastCell]:
            if self.right:
                attributes['borderright'] = self.right.value()


        return attributes



    ''' paragraph attributes
    '''
    def paragraph_attributes(self, nesting_level=0):
        attributes = {}

        # top and bottom
        if self.top:
            attributes['bordertop'] = self.top.value()

        if self.bottom:
            attributes['borderbottom'] = self.bottom.value()

        if self.left:
            attributes['borderleft'] = self.left.value()

        if self.right:
            attributes['borderright'] = self.right.value()

        return attributes



''' gsheet cell border object wrapper
'''
class Border(object):

    ''' constructor
    '''
    def __init__(self, border_dict, nesting_level=0):
        self.style = None
        self.width = None
        self.color = None

        if border_dict:
            self.width = border_dict.get('width') / ODT_BORDER_WIDTH_FACTOR
            self.color = RgbColor(border_dict.get('color'))

            # TODO: handle double
            self.style = GSHEET_ODT_BORDER_MAPPING.get(self.style, 'solid')


    ''' string representation
    '''
    def __repr__(self):
        return f"{self.width}pt {self.style} {self.color.value()}"


    ''' value
    '''
    def value(self, nesting_level=0):
        return f"{self.width}pt {self.style} {self.color.value()}"



''' Cell Merge spec wrapper
'''
class CellMergeSpec(object):
    def __init__(self, nesting_level=0):
        self.multi_col = MultiSpan.No
        self.multi_row = MultiSpan.No

        self.col_span = 1
        self.row_span = 1


    ''' string representation
    '''
    def to_string(self, nesting_level=0):
        return f"multicolumn: {self.multi_col}, multirow: {self.multi_row}"


    ''' table-cell attributes
    '''
    def table_cell_attributes(self, nesting_level=0):
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
    def __init__(self, row_metadata_dict, nesting_level=0):
        self.pixel_size = int(row_metadata_dict['pixelSize'])
        self.inches = row_height_in_inches(self.pixel_size)



''' gsheet columnMetadata object wrapper
'''
class ColumnMetadata(object):

    ''' constructor
    '''
    def __init__(self, column_metadata_dict, nesting_level=0):
        self.pixel_size = int(column_metadata_dict['pixelSize'])



''' gsheet merge object wrapper
'''
class Merge(object):

    ''' constructor
    '''
    def __init__(self, gsheet_merge_dict, start_row, start_column, nesting_level=0):
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
    def __init__(self, rgb_dict=None, nesting_level=0):
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
    def key(self, nesting_level=0):
        return ''.join('{:02x}'.format(a) for a in [self.red, self.green, self.blue])


    ''' color value for tabularray color
    '''
    def value(self, nesting_level=0):
        return '#' + ''.join('{:02x}'.format(a) for a in [self.red, self.green, self.blue])


    ''' is the color white
    '''
    def is_white(self, nesting_level=0):
        if self.red == 255 and self.green == 255 and self.blue == 255:
            return True

        return False



''' gsheet cell padding object wrapper
'''
class Padding(object):

    ''' constructor
    '''
    def __init__(self, padding_dict=None, nesting_level=0):
        # HACK: paddings are hard-coded
        if padding_dict:
            # self.top = int(padding_dict.get('top', 0))
            # self.right = int(padding_dict.get('right', 0))
            # self.bottom = int(padding_dict.get('bottom', 0))
            # self.left = int(padding_dict.get('left', 0))
            self.top = 1
            self.right = 2
            self.bottom = 1
            self.left = 2
        else:
            self.top = 1
            self.right = 2
            self.bottom = 1
            self.left = 2


    ''' string representation
    '''
    def table_cell_attributes(self, nesting_level=0):
        attributes = {}

        attributes['paddingtop'] = f"{self.top}pt"
        attributes['paddingright'] = f"{self.right}pt"
        attributes['paddingbottom'] = f"{self.bottom}pt"
        attributes['paddingleft'] = f"{self.left}pt"

        return attributes



''' gsheet text format run object wrapper
'''
class TextFormatRun(object):

    ''' constructor
    '''
    def __init__(self, run_dict=None, default_format=None, nesting_level=0):
        if run_dict:
            self.start_index = int(run_dict.get('startIndex', 0))
            format = run_dict.get('format')
            new_format = {**default_format, **format}
            self.format = TextFormat(new_format)
        else:
            self.start_index = None
            self.format = None


    ''' generates the odt code
    '''
    def text_attributes(self, text, nesting_level=0):
        text_attributes = self.format.text_attributes()

        return {'text': text[self.start_index:], 'text-attributes': text_attributes}



''' gsheet cell notes object wrapper
    TODO: handle keep-with-previous defined in notes
'''
class CellNote(object):

    ''' constructor
    '''
    def __init__(self, document_nesting_depth, note_dict={}, nesting_level=0):
        self.outline_level = 0
        self.free_content = False
        self.content = note_dict.get('content', None)
        self.style = note_dict.get('style', None)
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
                self.outline_level = outline_level_object['outline-level'] + document_nesting_depth
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


    ''' style attributes dict to create Style
    '''
    def style_attributes(self, nesting_level=0):
        attributes = {}

        if self.style is not None:
            attributes['parentstylename'] = self.style

        return attributes


    ''' paragraph attrubutes dict to craete ParagraphProperties
    '''
    def paragraph_attributes(self, nesting_level=0):
        attributes = {}

        if self.new_page:
            attributes['breakbefore'] = 'page'

        if self.keep_with_next:
            attributes['keepwithnext'] = 'always'

        return attributes



''' gsheet cell inline-image wrapper
'''
class InlineImage(object):

    ''' constructor
    '''
    def __init__(self, ii_dict={}, nesting_level=0):
        self.ii_dict = ii_dict
        self.file_path = ii_dict.get('file-path', None)
        self.file_type = ii_dict.get('file-type', None)
        self.image_width = ii_dict.get('image-width', None)
        self.image_height = ii_dict.get('image-height', None)
        self.type = ii_dict.get('type', 'inline')
        self.fit_height_to_container = ii_dict.get('fit-height-to-container', False)
        self.fit_width_to_container = ii_dict.get('fit-width-to-container', False)
        self.keep_aspect_ratio = ii_dict.get('keep-aspect-ratio', True)
        self.position = ii_dict.get('position', 'center center')
        self.wrap = ii_dict.get('wrap', 'parallel')

        self.anchor_type = 'paragraph'

        self.halign, self.valign = 'center', 'center'
        positions = self.position.split(' ')
        if len(positions) == 2:
            self.halign, self.valign = positions[0], positions[1]
        elif len(positions) == 1:
            self.halign = positions[0]

    
    ''' attributes dict for GraphicProperties
    '''
    def graphic_properties_attributes(self, nesting_level=0):
        attributes = {'verticalpos': self.valign, 'horizontalpos': self.halign, 'verticalrel': self.anchor_type, 'wrap': self.wrap, 'runthrough': 'background', 'flowwithtext': True}

        return attributes


    ''' attributes dict for DrawFrame
    '''
    def frame_attributes(self, container_width=None, container_height=None, preserve=None, nesting_level=0):
        width_in_inches, height_in_inches = self.image_width, self.image_height
        if self.fit_height_to_container or self.fit_width_to_container:
            if container_width and container_height:
                width_in_inches, height_in_inches, scale = fit_width_height(fit_within_width=container_width, fit_within_height=container_height, width_to_fit=self.image_width, height_to_fit=self.image_height)

        attributes = {'anchortype': self.anchor_type, 'width': f"{width_in_inches}in", 'height': f"{height_in_inches}in"}
        if self.fit_height_to_container:
            if self.keep_aspect_ratio:
                attributes = {**attributes, **{'relheight': '100%', 'relwidth': 'scale'}}
            else:
                attributes = {**attributes, **{'relheight': '100%', 'relwidth': 'scale-min'}}


        if self.fit_width_to_container:
            if self.keep_aspect_ratio:
                attributes = {**attributes, **{'relwidth': '100%', 'relheight': 'scale'}}
            else:
                attributes = {**attributes, **{'relwidth': '100%', 'relheight': 'scale-min'}}

        if preserve == 'height':
            attributes = {**attributes, **{'relheight': '100%', 'relwidth': 'scale-min'}}

        elif preserve == 'width':
            attributes = {**attributes, **{'relheight': 'scale-min', 'relwidth': '100%'}}

        return attributes



''' gsheet vertical alignment object wrapper
'''
class VerticalAlignment(object):

    ''' constructor
    '''
    def __init__(self, valign=None, nesting_level=0):
        if valign:
            self.valign = TEXT_VALIGN_MAP.get(valign, 'top')
        else:
            self.valign = TEXT_VALIGN_MAP.get('TOP')



''' gsheet horizontal alignment object wrapper
'''
class HorizontalAlignment(object):

    ''' constructor
    '''
    def __init__(self, halign=None, nesting_level=0):
        if halign:
            self.halign = TEXT_HALIGN_MAP.get(halign, 'left')
        else:
            self.halign = TEXT_HALIGN_MAP.get('LEFT')



''' gsheet wrapping object wrapper
'''
class Wrapping(object):

    ''' constructor
    '''
    def __init__(self, wrap=None, nesting_level=0):
        if wrap:
            self.wrapping = WRAP_STRATEGY_MAP.get(wrap, 'WRAP')
        else:
            self.wrapping = WRAP_STRATEGY_MAP.get('WRAP')



''' gsheet textRotation object wrapper
'''
class TextRotation(object):

    ''' constructor
    '''
    def __init__(self, text_rotation=None, nesting_level=0):
        self.angle = None

        if text_rotation:
            if 'vertical' in text_rotation and text_rotation['vertical'] == False:
                self.angle = 90

            if 'angle' in text_rotation:
                self.angle = text_rotation["angle"]



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
def process_table(odt, section_data, nesting_level=0):
    section = OdtTableSection(odt=odt, section_data=section_data, nesting_level=nesting_level)
    section.section_to_odt(nesting_level=nesting_level)



''' Gsheet processor
'''
def process_gsheet(odt, section_data, nesting_level=0):
    section = OdtGsheetSection(odt=odt, section_data=section_data, nesting_level=nesting_level)
    section.section_to_odt(nesting_level=nesting_level)



''' Table of Content processor
'''
def process_toc(odt, section_data, nesting_level=0):
    section = OdtToCSection(odt=odt, section_data=section_data, nesting_level=nesting_level)
    section.section_to_odt(nesting_level=nesting_level)



''' List of Figure processor
'''
def process_lof(odt, section_data, nesting_level=0):
    section = OdtLoFSection(odt=odt, section_data=section_data, nesting_level=nesting_level)
    section.section_to_odt(nesting_level=nesting_level)



''' List of Table processor
'''
def process_lot(odt, section_data, nesting_level=0):
    section = OdtLoTSection(odt=odt, section_data=section_data, nesting_level=nesting_level)
    section.section_to_odt(nesting_level=nesting_level)



''' pdf processor
'''
def process_pdf(odt, section_data, nesting_level=0):
    section = OdtPdfSection(odt=odt, section_data=section_data, nesting_level=nesting_level)
    section.section_to_odt(nesting_level=nesting_level)



''' odt processor
'''
def process_odt(odt, section_data, nesting_level=0):
    warn(f"content type [odt] not supported")



''' docx processor
'''
def process_docx(odt, section_data, nesting_level=0):
    warn(f"content type [docx] not supported")
