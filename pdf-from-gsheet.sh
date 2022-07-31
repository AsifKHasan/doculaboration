#!/usr/bin/env bash
# gsheet->json->pandoc->pdf pipeline

# parameters
DOCUMENT=$1

# set echo off

# json-from-gsheet
pushd ./gsheet-to-json/src
# ./json-from-gsheet.py --config "../conf/config.yml" --gsheet ${DOCUMENT}

if [ ${?} -ne 0 ]; then
  popd && exit 1
else
  popd
fi

# pandoc-from-json
pushd ./json-to-pandoc/src
# ./pandoc-from-json.py --config "../conf/config.yml" --json ${DOCUMENT}

if [ ${?} -ne 0 ]; then
  popd && exit 1
else
  popd
fi

# pandoc-from-json
pushd ./out
time pandoc ${DOCUMENT}.mkd ../json-to-pandoc/conf/preamble.yml -s --pdf-engine=lualatex -f markdown -t latex -o ${DOCUMENT}.pdf

if [ ${?} -ne 0 ];  then
  popd && exit 1
else
  popd
fi
