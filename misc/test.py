from pprint import pprint
from odf.opendocument import OpenDocumentText, load
from odf.style import MasterPage, PageLayout

# ODT_PATH = '../out/template-spectrum.odt'
ODT_PATH = '../out/NBR-NSW__functional-and-technical-report__PwC__v5.0.odt'

odt = load(ODT_PATH)
# odt = OpenDocumentText()

standard_master_page_name = 'Standard'
masterstyles = odt.masterstyles
standard_master_page = None
for master_page in masterstyles.getElementsByType(MasterPage):
    if master_page.getAttribute('name') == standard_master_page_name:
        standard_master_page = master_page
        break

if standard_master_page is not None:
    print(f"master-page {standard_master_page.getAttribute('name')} found; page-layout is {standard_master_page.attributes[(standard_master_page.qname[0], 'page-layout-name')]}")
    # print(standard_master_page.qname)
    # for k in standard_master_page.attributes.keys():
        # print(k[1], ":", standard_master_page.attributes[k])
        # print(standard_master_page.attributes[k])

    # pprint(odt.stylesxml())
else:
    pprint(f"master-page {standard_master_page_name} NOT found")


# for pl in odt.automaticstyles.getElementsByType(PageLayout):
# for pl in odt.masterstyles.childNodes:
#     pprint(pl.getAttribute('name'))

# odt.save(ODT_PATH)
