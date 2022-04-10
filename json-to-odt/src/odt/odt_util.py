#!/usr/bin/env python3

''' various utilities for generating an Openoffice odt document
'''

from odf.style import MasterPage, PageLayout, PageLayoutProperties
from odf.text import P

from helper.logger import *

''' write a paragraph in a given style
'''
def create_paragraph(odt, style_name, text):
    style = None
    if not style_name is None: 
        style = odt.getStyleByName(style_name)

    p = P(stylename=style, text=text)
    odt.text.addElement(p)


''' gets the page-layout from odt if it is there, else create one
'''
def get_or_create_page_layout(odt, odt_specs, page_layout_name, page_spec, margin_spec, orientation):
    page_layout = None

    for pl in odt.automaticstyles.getElementsByType(PageLayout):
        if pl.getAttribute('name') == page_layout_name:
            page_layout = pl
            break

    if page_layout is None:
        # create one
        page_layout = PageLayout(name=page_layout_name)
        odt.automaticstyles.addElement(page_layout)
        pageheight = f"{odt_specs['page-spec'][page_spec]['height']}in"
        pagewidth = f"{odt_specs['page-spec'][page_spec]['width']}in"
        margintop = f"{odt_specs['margin-spec'][margin_spec]['top']}in"
        marginbottom = f"{odt_specs['margin-spec'][margin_spec]['bottom']}in"
        marginleft = f"{odt_specs['margin-spec'][margin_spec]['left']}in"
        marginright = f"{odt_specs['margin-spec'][margin_spec]['right']}in"
        # margingutter = f"{odt_specs['margin-spec'][margin_spec]['gutter']}in"
        printorientation = orientation
        page_layout.addElement(PageLayoutProperties(pagewidth=pagewidth, pageheight=pageheight, margintop=margintop, marginbottom=marginbottom, marginleft=marginleft, marginright=marginright, printorientation=orientation))

    return page_layout


''' new odt master-page
    page layouts are saved with a name page-spec__margin-spec__[portrait|landscape]
'''
def get_or_create_master_page(odt, odt_specs, page_layout_name, page_spec, margin_spec, orientation):
    # see if the master-page already exists or not in the odt
    master_page = None

    for mp in odt.masterstyles.childNodes:
        if mp.getAttribute('name') == page_layout_name:
            master_page = mp
            break

    if master_page is None:
        # create one, first get/create the page-layout
        page_layout = get_or_create_page_layout(odt, odt_specs, page_layout_name, page_spec, margin_spec, orientation)
        master_page = MasterPage(name=page_layout_name, pagelayoutname=page_layout_name)
        odt.masterstyles.addElement(master_page)

    return master_page
