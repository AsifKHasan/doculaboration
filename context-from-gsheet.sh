#!/usr/bin/env bash
# gsheet->json->context->pdf pipeline

# parameters
DOCUMENT=$1

# set echo off
PYTHON=python3

# json-from-gsheet
pushd ./gsheet-to-json/src
${PYTHON} json-from-gsheet.py --config "../conf/config.yml" --gsheet ${DOCUMENT}

if [ ${?} -ne 0 ]; then
  popd && exit 1
else
  popd
fi

# context-from-json
pushd ./json-to-context/src
${PYTHON} context-from-json.py --config "../conf/config.yml" --json ${DOCUMENT}

if [ ${?} -ne 0 ]; then
  popd && exit 1
else
  popd
fi

# context -> pdf
pushd ./out
time context --run ${DOCUMENT}.context.tex

if [ ${?} -ne 0 ];  then
  popd && exit 1
else
  popd
fi
