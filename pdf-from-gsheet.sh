#!/usr/bin/env bash
# gsheet->json->latex->pdf pipeline

# parameters
DOCUMENT=$1

# set echo off
PYTHON=python3

# json-from-gsheet
pushd ./gsheet-to-json/src
# ${PYTHON} json-from-gsheet.py --config "../conf/config.yml" --gsheet ${DOCUMENT}

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

# latex-from-json
pushd ./out
time pandoc ${DOCUMENT}.mkd ../json-to-latex/conf/preamble.yml -s --pdf-engine=lualatex -f markdown -t latex -o ${DOCUMENT}.pdf

if [ ${?} -ne 0 ];  then
  popd && exit 1
else
  popd
fi
