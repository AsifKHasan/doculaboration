#!/usr/bin/env bash
# gsheet->json->docx pipeline

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

# docx-from-json
pushd ./json-to-docx/src
${PYTHON} docx-from-json.py --config "../conf/config.yml" --json ${DOCUMENT}

if [ ${?} -ne 0 ]; then
  popd && exit 1
else
  popd
fi
