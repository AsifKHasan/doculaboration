#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import odf
from odf.opendocument import OpenDocumentText
from odf.element import Element
from odf.text import P
from odf.math import Math
from namespaces import MATHNS


formula =u'<mrow><msup><mi>x</mi><mn>2</mn></msup><mo>&#x0002B;</mo><msup><mi>y</mi><mn>2</mn></msup><mo>&#x0003D;</mo><msup><mi>z</mi><mn>2</mn></msup></mrow>'

def main():
    doc = OpenDocumentText()
    p = P(text=u'text')
    df = odf.draw.Frame( zindex=0, anchortype='as-char')
    p.addElement(df)
    doc.text.addElement(p)

    math = Math()
    annot = Element(qname = (MATHNS,u'annotation'))
    annot.addText(formula, check_grammar=False)
    annot.setAttribute((MATHNS,'encoding'), 'StarMath 5.0', check_grammar=False)
    math.addElement(annot)
    do = odf.draw.Object()
    do.addElement(math)
    df.addElement(do)

    outputfile = u'test2'
    doc.save(outputfile, True)

if __name__ == '__main__':
    main()