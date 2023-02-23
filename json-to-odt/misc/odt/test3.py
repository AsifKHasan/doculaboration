#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import latex2mathml.converter

latex_input = '$a_{1..n} ~ and ~ b^{1..n}$'
mathml_output = latex2mathml.converter.convert(latex_input)
print(mathml_output)