#!/usr/bin/env bash
# gsheet->json->latex->pdf pipeline

# parameters
DOCUMENT=$1

# set echo off
PYTHON=python

# json-from-gsheet
pushd ./gsheet-to-json/src
${PYTHON} json-from-gsheet.py --config "../conf/config.yml" --gsheet ${DOCUMENT}

if [ ${?} -ne 0 ]; then
  popd && exit 1
else
  popd
fi

# latex-from-json
pushd ./json-to-latex/src
${PYTHON} latex-from-json.py --config "../conf/config.yml" --json ${DOCUMENT}

if [ ${?} -ne 0 ]; then
  popd && exit 1
else
  popd
fi

# latex -> pdf
pushd ./out
time lualatex ${DOCUMENT}.latex.tex --output-format=pdf

if [ ${?} -ne 0 ];  then
  popd && exit 1
else
  popd
fi
