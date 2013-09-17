#!/bin/bash

if [[ -e $1 ]]; then 
	bin/xls2csv.njs --read $1
	exit
fi

for i in test_files/*.xls; do 
	bin/xls2csv.njs --read $i 2>&1 >/dev/null | sed 's/^.*parsing //'
done
