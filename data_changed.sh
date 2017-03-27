#!/usr/bin/env bash

ls -l resources/data.xlsx

echo docker-compose stop
docker-compose stop

echo rm -rf resources/generated_notes/\*
rm -rf resources/generated_notes/*

echo node prepare_gennerated.js
node prepare_gennerated.js > log_`date "+%Y-%m-%d_%H-%M-%S"`.txt 2>&1 

echo docker-compose up -d
docker-compose up -d
