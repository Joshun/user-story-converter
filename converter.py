import sys
import csv
from docx import Document
from docx.shared import Inches

document = Document()
document.add_heading("User Stories")

table = document.add_table(rows=1, cols=3)
header_rows = table.rows[0].cells
header_rows[0].text = "#"
header_rows[1].text = "Story"
header_rows[2].text = "Acceptance criteria"

def show_help():
    print("Usage: converter.py <stories.csv>")
        

if "-h" in sys.argv:
    show_help()
    sys.exit(0)

elif len(sys.argv) != 2:
    show_help()
    sys.exit(1)


infile_name = sys.argv[1]

with open(infile_name) as f:
    stories_reader = csv.reader(f)
    for i, row in enumerate(stories_reader):

        # Don't include csv header
        if i==0:
            continue

        tag, num, story, points, add_info, excluded = row
        acc_criteria = add_info.split("\n")

        row_cells = table.add_row().cells
        row_cells[0].text = str(num)
        row_cells[1].text = str(story)
        # row_cells[2].text = add_info
        for a in acc_criteria:
            if len(a) < 2:
                continue
            # pg.add_run(a)
            pg = row_cells[2].add_paragraph(a, style="List Bullet")


if infile_name[-4:] == ".csv":
    outfile_name = infile_name[:-4] + ".docx"

document.save(outfile_name)
print("Written to " + outfile_name)