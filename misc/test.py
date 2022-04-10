from pprint import pprint
from odf.opendocument import OpenDocumentText, load
from odf.style import PageLayout

ODT_PATH = '../out/template-spectrum.odt'
# ODT_PATH = '../out/NBR-NSW__functional-and-technical-report__PwC__v5.0.odt'

# odt = load(ODT_PATH)


odt = OpenDocumentText()


# masterstyles = odt.masterstyles
# pprint(ET.tostring(masterstyles))

# for pl in odt.automaticstyles.getElementsByType(PageLayout):
for pl in odt.masterstyles.childNodes:
    pprint(pl.getAttribute('name'))

odt.save(ODT_PATH)
