#!/bin/sh
# assuming the script is in the program directory
SRC="$(dirname "$0")"/main.py 
PYTHON=/usr/bin/python3
$PYTHON $SRC "$@"
