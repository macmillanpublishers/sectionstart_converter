#! /bin/bash

read -p " * * * Update validated files for tests now? (y or n) * * *" answer

while true
do
  case $answer in
   [yY]* ) cd "$( dirname "${BASH_SOURCE[0]}" )" && python ./rsvalidate_transform_tests.py update_valid_outputs
           break;;

   * ) echo "Okay, exiting."
   		   exit;;
  esac
done
