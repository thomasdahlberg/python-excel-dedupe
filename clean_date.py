import sys
import csv

file_name = sys.argv[1]

tsv_file = open(file_name)

read_tsv = csv.reader(tsv_file, delimiter="\t")

for row in read_tsv:
    print(row)
