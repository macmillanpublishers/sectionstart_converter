#! /bin/bash

read -p " * * * Update validated files for ravalidate transform tests now? (y or n) * * *" answer

while true
do
  case $answer in
   [yY]* ) cd "$( dirname "${BASH_SOURCE[0]}" )" && python ./rsvalidate_transform_tests.py update_valid_outputs
           break;;

   * ) echo "Okay..."
          break;;
  esac
done

read -p " * * * Update validated files for isbncheck transform tests now? (y or n) * * *" answer_b

while true
do
  case $answer_b in
   [yY]* ) cd "$( dirname "${BASH_SOURCE[0]}" )" && python ./isbncheck_transform_tests.py update_valid_outputs
           break;;

   * ) echo "Okay, exiting."
   		   exit;;
  esac
done
