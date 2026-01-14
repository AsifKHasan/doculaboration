#!/usr/bin/env python

from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

A4_W = Inches(8.27)
A4_H = Inches(11.69)

def set_a4_no_margins(section):
    section.page_width = A4_W
    section.page_height = A4_H
    section.top_margin = Inches(0)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0)
    section.right_margin = Inches(0)
    section.header_distance = Inches(0)
    section.footer_distance = Inches(0)

def inline_to_anchored_behind(inline):
    """
    Convert an inline picture (<wp:inline>) to a floating anchored picture (<wp:anchor>)
    positioned at page origin and behind text.
    """
    inline_elm = inline._inline  # CT_Inline element

    anchor = OxmlElement("wp:anchor")
    anchor.set("simplePos", "0")
    anchor.set("relativeHeight", "0")
    anchor.set("behindDoc", "1")          # behind text
    anchor.set("locked", "0")
    anchor.set("layoutInCell", "1")
    anchor.set("allowOverlap", "1")
    anchor.set("distT", "0")
    anchor.set("distB", "0")
    anchor.set("distL", "0")
    anchor.set("distR", "0")

    # required children: wp:simplePos, wp:positionH, wp:positionV, wp:extent, wp:effectExtent, wp:wrapNone, ...
    simplePos = OxmlElement("wp:simplePos")
    simplePos.set("x", "0")
    simplePos.set("y", "0")
    anchor.append(simplePos)

    positionH = OxmlElement("wp:positionH")
    positionH.set("relativeFrom", "page")
    posOffsetH = OxmlElement("wp:posOffset")
    posOffsetH.text = "0"
    positionH.append(posOffsetH)
    anchor.append(positionH)

    positionV = OxmlElement("wp:positionV")
    positionV.set("relativeFrom", "page")
    posOffsetV = OxmlElement("wp:posOffset")
    posOffsetV.text = "0"
    positionV.append(posOffsetV)
    anchor.append(positionV)

    # Move over the existing wp:extent (size) and other needed children from inline to anchor
    # We keep the same graphic and size Word generated.
    for child in list(inline_elm):
        # We'll skip wp:docPr because anchor needs it too (we can keep it)
        anchor.append(child)

    wrapNone = OxmlElement("wp:wrapNone")
    anchor.append(wrapNone)

    # Replace inline with anchor in the parent <w:drawing>
    drawing = inline_elm.getparent()
    drawing.remove(inline_elm)
    drawing.append(anchor)

def add_background_image_to_header(header, image_path, width=A4_W, height=A4_H):
    """
    Inserts an image into the header, stretches it to page size, and makes it floating behind text.
    """
    # Put it in its own paragraph at start
    p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    run = p.add_run()
    inline = run.add_picture(image_path, width=width, height=height)
    inline_to_anchored_behind(inline)

def clear_header_footer(hf):
    # Remove all paragraphs in header/footer so we can set our own cleanly
    for p in list(hf.paragraphs):
        p._element.getparent().remove(p._element)

def build_first_page_header_footer(section):
    """
    Customize page 1 header/footer here.
    """
    # FIRST PAGE header/footer objects
    h = section.first_page_header
    f = section.first_page_footer

    clear_header_footer(h)
    clear_header_footer(f)

    # Example content (replace with yours)
    h.add_paragraph("FIRST PAGE HEADER")
    f.add_paragraph("FIRST PAGE FOOTER")

def build_odd_header_footer(section):
    """
    Customize ODD (primary) header/footer template here.
    """
    h = section.header
    f = section.footer

    clear_header_footer(h)
    clear_header_footer(f)

    h.add_paragraph("ODD PAGE HEADER")
    f.add_paragraph("ODD PAGE FOOTER")

def build_even_header_footer(section):
    """
    Customize EVEN header/footer template here.
    """
    h = section.even_page_header
    f = section.even_page_footer

    clear_header_footer(h)
    clear_header_footer(f)

    h.add_paragraph("EVEN PAGE HEADER")
    f.add_paragraph("EVEN PAGE FOOTER")

def create_doc_with_5_background_pages(
    image_paths,  # list of 5 A4 images
    output_path="output.docx",
):
    assert len(image_paths) == 5, "Need exactly 5 image paths."

    doc = Document()

    # Global setting: enable different odd/even headers across the document
    doc.settings.odd_and_even_pages_header_footer = True

    # -------- Section 1 (Page 1) --------
    s1 = doc.sections[0]
    set_a4_no_margins(s1)

    # Enable "Different first page" for the first section so page 1 can use first_page_header/footer
    s1.different_first_page_header_footer = True

    # Build content for first page header/footer
    build_first_page_header_footer(s1)

    # Put background image for page 1 in first-page header
    add_background_image_to_header(s1.first_page_header, image_paths[0])

    # Ensure page 1 has at least one body paragraph so the page exists
    doc.add_paragraph("")

    # -------- Sections 2..5 (Pages 2..5) --------
    # These pages should use odd/even header/footer. We'll build both for each section,
    # then inject the background into the header that matches the page parity.
    for i in range(2, 6):  # page numbers 2..5
        sec = doc.add_section(WD_SECTION.NEW_PAGE)
        set_a4_no_margins(sec)

        # For these sections, we do NOT need different_first_page_header_footer
        sec.different_first_page_header_footer = False

        # Make sure each section does NOT link headers/footers to previous,
        # otherwise the background images would get inherited.
        sec.header.is_linked_to_previous = False
        sec.footer.is_linked_to_previous = False
        sec.even_page_header.is_linked_to_previous = False
        sec.even_page_footer.is_linked_to_previous = False

        # Build odd/even templates
        build_odd_header_footer(sec)
        build_even_header_footer(sec)

        # Add the background image to the header Word will use for that page
        img = image_paths[i - 1]  # image index 1..4
        if i % 2 == 0:
            add_background_image_to_header(sec.even_page_header, img)
        else:
            add_background_image_to_header(sec.header, img)

        # body content so the page exists
        doc.add_paragraph("")

    doc.save(output_path)
    return output_path

if __name__ == "__main__":
    images = [
        "../../out/tmp/a4-page-bg-01.jpg",
        "../../out/tmp/3.__JV Agreement-a__0001-1.jpg",
        "../../out/tmp/3.__JV Agreement-a__0001-2.jpg",
        "../../out/tmp/3.__JV Agreement-a__0001-3.jpg",
        "../../out/tmp/3.__JV Agreement-a__0001-4.jpg",
    ]
    create_doc_with_5_background_pages(images, "five_pages_backgrounds.docx")
