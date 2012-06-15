#!/bin/bash


script="./misc.awk"
ppedir="../miscellaneous-expenses/"
#fordir=${ppedir}"formatted/"
fordir=${ppedir}

# $ ./misc.awk fixed=1 ../miscellaneous-expenses/2011.txt > ../miscellaneous-expenses/formatted

function help_message {
  echo
  echo This script supports a build from scratch of an LTIB software project using source code from the GEI
  echo Git revision control system.  The script will build in the current working directory.
  echo
  echo usage:
  echo
  echo "$0 [-h] [-e] [-f]"
  echo
  echo "-h: display this help message and exit"
  echo "-e: echo the command that would be executed and exit"
  echo "-f: force all files to be written even if they already exist"
  echo
  echo examples:
  echo
  echo "$ $0"
  echo "$ $0 -f"
  echo "$ $0 -e"
  echo "$ $0 -ef"
  echo 
  exit 0
}

function error_message {
  echo "$0 -h # pass the -h flag to see the help options"
  exit 1
}

force=0
doecho=0
while getopts "efh" opt; do
  case $opt in
  e) # local repository
    doecho=1
    ;;
  f) # local repository
    force=1
    ;;
  h) # help
    help_message
    ;;
  *) # unknown
    error_message
    ;;
  esac
done

if ! [ -x ${script} ]; then
  echo "error: ${script} either does not exist or is not executable"
  exit 1
fi 
if ! [ -d ${ppedir} ]; then
  echo "error: ${ppedir} is not a directory"
  exit 1
fi 
if ! [ -d ${fordir} ]; then
  echo "error: ${fordir} is not a directory"
  exit 1
fi 

for file in ${ppedir}*.txt; do
  bn=`basename ${file} .txt`
  if [ ${force} -eq 1 ] || ! [ -f ${fordir}${bn}.csv ]; then
    newcommand=`echo "${script} -v print_header=1 fixed=1 ${file} > ${fordir}${bn}.csv; "`
    if [ ${doecho} -eq 1 ]; then
      echo ${newcommand}
    fi
    mycommand=${mycommand}${newcommand}
  fi
done

if [ ${doecho} -eq 0 ]; then
  eval ${mycommand}
fi
