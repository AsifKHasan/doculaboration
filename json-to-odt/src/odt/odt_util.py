#!/usr/bin/env python3

''' various utilities for generating an Openoffice odt document
'''
import platform
import subprocess
from pathlib import Path

from odf.style import Style, MasterPage, PageLayout, PageLayoutProperties, TextProperties, ParagraphProperties
from odf.text import P, TableOfContent, TableOfContentSource, IndexTitleTemplate

from helper.logger import *


if platform.system() == 'Windows':
    LIBREOFFICE_EXECUTABLE = 'C:/Program Files/LibreOffice/program/soffice.exe'
else:
    LIBREOFFICE_EXECUTABLE = 'soffice'


# --------------------------------------------------------------------------------------------------------------------------------------------
# paragraphs and styles

''' create style - family paragraph
'''
def create_paragraph_style(odt, style_name, parent_style_name, page_break=False, master_page_name=None):
    if master_page_name is not None:
        paragraph_style = Style(name=style_name, parentstylename=parent_style_name, family="paragraph", masterpagename=master_page_name)
        odt.automaticstyles.addElement(paragraph_style)
        return

    if page_break:
        paragraph_style = Style(name=style_name, parentstylename=parent_style_name, family="paragraph")
        paragraph_style.addElement(ParagraphProperties(breakbefore="page"))
        odt.automaticstyles.addElement(paragraph_style)
        return
    else:
        paragraph_style = Style(name=style_name, parentstylename=parent_style_name, family="paragraph")
        odt.automaticstyles.addElement(paragraph_style)
        return


''' write a paragraph in a given style
'''
def create_paragraph(odt, style_name, text):
    style = odt.getStyleByName(style_name)
    p = P(stylename=style, text=text)
    odt.text.addElement(p)



# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# indexes and pdf generation

''' update indexes through a macro macro:///Standard.Module1.open_document(document_url) which must be in OpenOffice macro library
'''
def update_indexes(odt, odt_path):
    document_url = Path(odt_path).as_uri()
    macro = f'"macro:///Standard.Module1.open_document("{document_url}")"'
    command_line = f'"{LIBREOFFICE_EXECUTABLE}" --headless --invisible {macro}'
    debug(f"updating indexes for {odt_path}")
    subprocess.call(command_line, shell=True);


''' given an odt file generates pdf in the given directory
'''
def generate_pdf(infile, outdir):
    command_line = f'"{LIBREOFFICE_EXECUTABLE}" --headless --convert-to pdf --outdir "{outdir}" "{infile}"'
    debug(f"generating pdf from {infile}")
    subprocess.call(command_line, shell=True);


''' create table-of-contents
'''
def create_toc(odt):
    name = 'Table of Content'
    toc = TableOfContent(name=name)
    toc_source = TableOfContentSource(outlinelevel=10)
    toc_title_template = IndexTitleTemplate()
    toc_source.addElement(toc_title_template)
    toc.addElement(toc_source)
    odt.text.addElement(toc)


''' create illustration-index
'''
def create_lof(odt):
    name = 'List of Figures'


''' create Table-index
'''
def create_lot(odt):
    name = 'List of Tables'



# -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# master-page and page-layout

''' update page-layout of Standard master-page with the given page-layout
'''
def update_standard_master_page(odt, new_page_layout):
    # get the Standard master-page
    standard_master_page_name = 'Standard'
    standard_master_page = None
    for master_page in odt.masterstyles.getElementsByType(MasterPage):
        if master_page.getAttribute('name') == standard_master_page_name:
            standard_master_page = master_page
            break

    if standard_master_page is not None:
        debug(f"master-page {standard_master_page.getAttribute('name')} found; page-layout is {standard_master_page.attributes[(standard_master_page.qname[0], 'page-layout-name')]}")
        standard_master_page.attributes[(standard_master_page.qname[0], 'page-layout-name')] = new_page_layout
        debug(f"master-page {standard_master_page.getAttribute('name')}  page-layout changed to {standard_master_page.attributes[(standard_master_page.qname[0], 'page-layout-name')]}")
    else:
        warn(f"master-page {standard_master_page_name} NOT found")


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
