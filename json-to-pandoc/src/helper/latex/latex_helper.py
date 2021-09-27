#!/usr/bin/env python3

''' Latex Table object wrapper
'''
class LatexTable(object):

    ''' constructor
    '''
    def __init__(self, row_data_list):
        pass


''' Latex section object wrapper
'''
class LatexSection(object):

    ''' constructor
    '''
    def __init__(self, section_data, section_break):
        section_contents = section_data.get('contents')

        self._section_break = section_break

        self._row_count = 0
        self._column_count = 0

        self._start_row = 0
        self._start_column = 0

        self._cells = []
        self._row_metadata_list = []
        self._column_metadata_list = []
        self._merge_list = []

        if section_contents:
            self._has_content = True

            sheets = section_contents.get('sheets')
            if isinstance(sheets, list) and len(sheets) > 0:
                properties = sheets[0].get('properties')
                if properties and 'gridProperties' in properties:
                    self._row_count = max(int(properties['gridProperties'].get('rowCount'), 0) - 2, 0)
                    self._column_count = max(int(properties['gridProperties'].get('columnCount'), 0) - 1, 0)

                data_list = sheets[0].get('data')
                if isinstance(data_list, list) and len(data_list) > 0:
                    data = data_list[0]
                    self._start_row = int(data.get('', 0))
                    self._start_column = int(data.get('', 0))

                    # rowData
                    for row_data in data.get('rowData', []):
                        self._cells.append(Row(row_data))

                    # rowMetaData
                    for row_metadata in data.get('rowMetaData', []):
                        self._row_metadata_list.append(RowMetaData(row_metadata))

                    # columnMetaData
                    for column_metadata in data.get('columnMetaData', []):
                        self._column_metadata_list.append(ColumnMetaData(column_metadata))

                    # merges
                    for merge in sheets[0].get('merges', []):
                        self._merge_list.append(Merge(merge, self._start_row, self._start_column))

        else:
            self._has_content = False


''' gsheet Cell object wrapper
'''
class Cell(object):

    ''' constructor
    '''
    def __init__(self, value):
        # TODO
        pass


''' gsheet Row object wrapper
'''
class Row(object):

    ''' constructor
    '''
    def __init__(self, row_data):
        self._cells = []
        for value in row_data.get('values', []):
            self._cells.append(Cell(value))


''' gsheet rowMetaData object wrapper
'''
class RowMetaData(object):

    ''' constructor
    '''
    def __init__(self, row_metadata_dict):
        self._pixel_size = int(row_metadata_dict['pixelSize'])


''' gsheet columnMetaData object wrapper
'''
class ColumnMetaData(object):

    ''' constructor
    '''
    def __init__(self, column_metadata_dict):
        self._pixel_size = int(column_metadata_dict['pixelSize'])


''' gsheet merge object wrapper
'''
class Merge(object):

    ''' constructor
    '''
    def __init__(self, gsheet_merge_dict, start_row, start_column):
        self._start_row = int(gsheet_merge_dict['startRowIndex']) - start_row
        self._end_row = int(gsheet_merge_dict['endRowIndex']) - start_row
        self._start_column = int(gsheet_merge_dict['startColumnIndex']) - start_column
        self._end_column = int(gsheet_merge_dict['endColumnIndex']) - start_column

        self._row_span = self._end_row - self._start_row
        self._column_span = self._end_column - self._start_column
