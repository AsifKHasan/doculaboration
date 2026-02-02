#!/usr/bin/env python
'''
'''

import yaml
from pathlib import Path

from helper.logger import *
from helper import logger

class ConfigService:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(ConfigService, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self, config_file=None, nesting_level=0):
        if self._initialized:
            return
        
        self._config_path = Path(config_file).resolve()
        self._config_dir = self._config_path.parent
        _config_dict = yaml.safe_load(open(self._config_path, 'r', encoding='utf-8'))

        self._log_level = _config_dict.get("log-level", 0)
        logger.LOG_LEVEL = self._log_level

        self._google_cred_json_path = Path(_config_dict.get('google-cred', None)).resolve()
        self._json_list = _config_dict.get('jsons', [])
        self._data_dir = Path(_config_dict.get('data-dir', '../data')).resolve()
        self._output_dir = Path(_config_dict.get('output-dir', '../../out')).resolve()

        self._temp_dir = self._output_dir / 'tmp'
        self._temp_dir.mkdir(parents=True, exist_ok=True)

        self._docx_template = Path(_config_dict.get('docx-template', None)).resolve()
        self._generate_pdf = _config_dict.get('generate-pdf', True)

        # page specs
        page_spec_file = self._config_dir / 'page-specs.yml'
        self._page_specs = yaml.safe_load(open(page_spec_file, 'r', encoding='utf-8'))

        # font specs
        # font_spec_file = self._config_dir / 'font-specs.yml'
        # if Path.exists(font_spec_file):
        #     self._font_specs = yaml.safe_load(open(font_spec_file, 'r', encoding='utf-8'))
        # else:
        #     warn(f"No font-spec [{font_spec_file}]' found .. no fonts to register", nesting_level=nesting_level)

        # style specs
        style_spec_file = self._config_dir / 'style-specs.yml'
        if Path.exists(page_spec_file):
            style_specs_original = yaml.safe_load(open(style_spec_file, 'r', encoding='utf-8'))

            # now we re-code the style-specs to something docx requires
            self._style_specs = {}
            for k, v in style_specs_original.items():
                self._style_specs[k] = self.recode_style_specs(style_specs_original=v, nesting_level=nesting_level)

        else:
            warn(f"No style-spec [{style_spec_file}]' found .. will not override any style", nesting_level=nesting_level)


        self._initialized = True


    ''' re-code the style-specs to something docx requires
    '''
    def recode_style_specs(self, style_specs_original, nesting_level=0):
        style_specs = {
            'ParagraphStyle': {
                'font': {
                    'name':               None,
                    'size':               None,
                    'color':              None,
                    'bold':               None,
                    'italic':             None,
                    'underline':          None,
                    'strike':             None,
                    'double_strike':      None,
                    'highlight_color':    None,
                    'all_caps':           None,
                    'small_caps':         None,
                    'subscript':          None,
                    'superscript':        None,
                    'complex_script':     None,
                    'cs_bold':            None,
                    'cs_italic':          None,
                    'emboss':             None,
                    'imprint':            None,
                    'outline':            None,
                    'shadow':             None,
                },

                'paragraph-format': {
                    'alignment':          None,

                    # margin
                    'left_indent':        None,
                    'right_indent':       None,
                    'space_before':       None,
                    'space_after':        None,
                }            
            },
            # shading
            'backgroundcolor':            None,

            # vertical alignment
            'verticalalign':              None,

            # border (sz val color space [shadow]]), val in dotted/dashed/single/thick/triple/double/none
            'borders': {
                'start':           None,
                'top':             None,
                'end':             None,
                'bottom':          None,
            },

            # padding
            'padding': {
                'start':          None,
                'top':            None,
                'end':            None,
                'bottom':         None,
            },

            'page-background': style_specs_original.get('page-background', []),
            'inline-image': style_specs_original.get('inline-image', []),
        }

        # re-code 'text-properties'
        if 'text-properties' in style_specs_original:
            text_properties = style_specs_original['text-properties']
            style_specs['ParagraphStyle']['font']['color'] = text_properties.get('color', None)
            style_specs['ParagraphStyle']['font']['name'] = text_properties.get('font.family.fontfamily', None)
            style_specs['ParagraphStyle']['font']['size'] = text_properties.get('font.size.fontsize', None)

            fontstyle = text_properties.get('font.style.fontstyle', None)
            if fontstyle == 'italic':
                style_specs['ParagraphStyle']['font']['italic'] = True

            fontweight = text_properties.get('font.weight.fontweight', None)
            if fontweight == 'bold':
                style_specs['ParagraphStyle']['font']['bold'] = True

        # re-code 'paragraph-properties'
        if 'paragraph-properties' in style_specs_original:
            paragraph_properties = style_specs_original['paragraph-properties']
            style_specs['ParagraphStyle']['paragraph-format']['alignment'] = paragraph_properties.get('textalign', None).upper()
            style_specs['verticalalign'] = paragraph_properties.get('verticalalign', None).upper()
            style_specs['backgroundcolor'] = paragraph_properties.get('backgroundcolor', None)

            # border
            style_specs['borders']['start'] = paragraph_properties.get('border.borderleft', None)
            style_specs['borders']['top'] = paragraph_properties.get('border.bordertop', None)
            style_specs['borders']['end'] = paragraph_properties.get('border.borderright', None)
            style_specs['borders']['bottom'] = paragraph_properties.get('border.borderbottom', None)

            # margin
            style_specs['ParagraphStyle']['paragraph-format']['left_indent'] = paragraph_properties.get('margin.marginleft', None)
            style_specs['ParagraphStyle']['paragraph-format']['space_before'] = paragraph_properties.get('margin.margintop', None)
            style_specs['ParagraphStyle']['paragraph-format']['right_indent'] = paragraph_properties.get('margin.marginright', None)
            style_specs['ParagraphStyle']['paragraph-format']['space_after'] = paragraph_properties.get('margin.marginbottom', None)

            # padding
            style_specs['padding']['start'] = paragraph_properties.get('padding.paddingleft', None)
            style_specs['padding']['top'] = paragraph_properties.get('padding.paddingtop', None)
            style_specs['padding']['end'] = paragraph_properties.get('padding.paddingright', None)
            style_specs['padding']['bottom'] = paragraph_properties.get('padding.paddingbottom', None)


        return style_specs