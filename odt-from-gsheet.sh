#!/usr/bin/env bash
# gsheet->json->odt pipeline

set echo off

# parameters
DOCUMENT=$1

# json-from-gsheet
pushd ./gsheet-to-json/src
./json-from-gsheet.py --config "../conf/config.yml" --gsheet ${DOCUMENT}

if [ ${?} -ne 0 ]; then
  popd && exit 1
else
  popd
fi

# odt-from-json
pushd ./json-to-odt/src
./odt-from-json.py --config "../conf/config.yml" --json ${DOCUMENT}

if [ ${?} -ne 0 ]; then
  popd && exit 1
else
  popd
fi
