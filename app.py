from xlsxcore import *

import glob

count = 0
xlsx_files = []
for file in glob.glob("*.xlsx"):
    xlsx_files.append(file)
    count += 1
    print("Loaded file number", count, file)

i = 0

for xlsx in xlsx_files:
    out_stream = xlsx2html(xlsx)
    out_stream.seek(0)

    file_name = xlsx[:-5]
    f = open(file_name + ".html", "w")
    f.write(out_stream.read())
    f.close()
    i += 1
    print("Converted file number", i, "of", count, file_name)