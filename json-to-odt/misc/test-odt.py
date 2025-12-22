from odf.opendocument import OpenDocumentText
from odf.style import PageLayout, PageLayoutProperties, MasterPage, Header, Footer
from odf.text import P, H, Section
from odf.namespaces import STYLENS

def make_master(doc, master_name, page_layout, header_text, footer_text, next_master=None):
    mp = MasterPage(name=master_name, pagelayoutname=page_layout)
    if next_master:
        mp.setAttribute("nextstylename", next_master)

    hdr = Header()
    hdr.addElement(P(text=header_text))
    mp.addElement(hdr)

    ftr = Footer()
    ftr.addElement(P(text=footer_text))
    mp.addElement(ftr)

    doc.masterstyles.addElement(mp)
    return mp

def make_layout(doc, layout_name, page_usage="all"):
    """
    page_usage: "all", "mirrored", "right", "left"
    """
    pl = PageLayout(name=layout_name)

    plp = PageLayoutProperties()
    if page_usage:
        # ODF attribute is style:page-usage (note the hyphen!)
        plp.setAttrNS(STYLENS, "page-usage", page_usage)

    pl.addElement(plp)
    doc.automaticstyles.addElement(pl)
    return pl

# --- Build document ---
doc = OpenDocumentText()

# Page layouts
first_layout = make_layout(doc, "SectionFirstLayout", page_usage="all")
right_layout = make_layout(doc, "SectionRightLayout", page_usage="right")
left_layout  = make_layout(doc, "SectionLeftLayout",  page_usage="left")

# Master pages:
# 1) First page master → then go to Right master
# 2) Right master → then go to Left master
# 3) Left master → then go to Right master (alternating)
make_master(
    doc,
    master_name="SectionFirstMaster",
    page_layout=first_layout,
    header_text="SECTION — FIRST PAGE HEADER",
    footer_text="SECTION — FIRST PAGE FOOTER",
    next_master="SectionRightMaster"
)

make_master(
    doc,
    master_name="SectionRightMaster",
    page_layout=right_layout,
    header_text="SECTION — ODD PAGE HEADER (Right pages)",
    footer_text="SECTION — ODD PAGE FOOTER (Right pages)",
    next_master="SectionLeftMaster"
)

make_master(
    doc,
    master_name="SectionLeftMaster",
    page_layout=left_layout,
    header_text="SECTION — EVEN PAGE HEADER (Left pages)",
    footer_text="SECTION — EVEN PAGE FOOTER (Left pages)",
    next_master="SectionRightMaster"
)

# --- Content before section (optional) ---
doc.text.addElement(H(outlinelevel=1, text="Document Title"))
doc.text.addElement(P(text="This is content before the special section."))

# --- The section that uses the special first/odd/even headers ---
sec = Section(name="MySpecialSection", masterpagename="SectionFirstMaster")
sec.addElement(H(outlinelevel=2, text="My Special Section"))

# Add enough content to force multiple pages.
# (How many paragraphs you need depends on default page size/margins.)
for i in range(1, 200):
    sec.addElement(P(text=f"Section paragraph {i}: Lorem ipsum dolor sit amet, consectetur adipiscing elit."))

doc.text.addElement(sec)

# --- Content after section (optional) ---
doc.text.addElement(P(text="This is content after the section."))

# Save
doc.save("section_first_odd_even_headers.odt")
print("Wrote section_first_odd_even_headers.odt")
