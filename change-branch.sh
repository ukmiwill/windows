#!/bin/bash

## declare an array variable

cd ~/tachograph-product-workspace/

arrRepo=(
tacho-driver-card
ui-contracts
tacho-web-app
tacho-sftp-s3
tacho-qe-gems
tacho-migration-service
tacho-internal-ui
tacho-external-ui
tacho-driver-internal-ui
tacho-driver-external-ui
tacho-company-card
tacho-build-dependencies
osl-data-dictionary
common-reminder-service

  "common-reminder-service"
  "osl-data-dictionary"
  "tacho-build-dependencies"
  "tacho-company-card"
  "tacho-driver-external-ui"
  "tacho-driver-internal-ui"
  "tacho-external-ui"
  "tacho-internal-ui"
  "tacho-migration-service"
  "tacho-qe-gems"
  "tacho-sftp-s3"
  "tacho-web-app"
  "ui-contracts"
)

arrBranch=(
  "main"
  "master"
  "master"
  "develop"
  "main"
  "main"
  "develop"
  "develop"
  "master"
  "master"
  "master"
  "master"
  "master"
)

for i in "${!arrRepo[@]}"
do
  cd ${arrRepo[i]}
  echo $PWD
  git checkout ${arrBranch[i]}
  cd ~/tachograph-product-workspace/
done
