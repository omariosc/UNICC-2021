import sys
import xlsxwriter
import pyexcel as pe

# Writes headings to xlsx workbook
def write_headings(worksheet):
  # extra_headings = ["FOLIO DDOCNAME", "Content Linked with folio in OLD PROD", "Total No. of Contents Linked with folio in OLD PROD", "Content Linked with folio in QA", "Total No. of Contents Linked with folio in QA", "Content Linked with folio in NEW PROD", "Total No. of Contents Linked with folio in NEW PROD"]
  headings = ["FOLIO DDOCNAME", "Content Linked with folio in OLD PROD", "Content Linked with folio in QA", "Content Linked with folio in NEW PROD"]
  col = 0
  for heading in headings:
    worksheet.write(0, col, heading)
    col += 1

def main():
  # Gets OLDPRD, QA and NEWPRD files from command line arguments
  old_prd_data_path = new_prd_data_path = qa_data_path = ""
  try:
    old_prd_data_path = sys.argv[1]
    new_prd_data_path = sys.argv[2]
    qa_data_path = sys.argv[3]
  except IndexError:
    print("Usage: python folio_analysis.py <OLDPRD DATA FILEPATH> <NEWPRD DATA FILEPATH> <QA DATA FILEPATH>")
    sys.exit(1)
  # Creates xlsx workbook and worksheet and writes headings
  workbook = xlsxwriter.Workbook((old_prd_data_path.replace("OLDPRD", "")).replace(".xlsx", "ANALYSIS.xlsx"))
  worksheet = workbook.add_worksheet("FolioAnalysis")
  write_headings(worksheet)
  # Sets initial row for worksheet
  row = 1
  qa_row = new_prd_row = 0
  # Reads content OLDPRD and NEWPRD data from xlsx workbook
  old_prd_content = pe.get_array(file_name=old_prd_data_path)[1:]
  new_prd_content = pe.get_array(file_name=new_prd_data_path)[1:]
  qa_content = pe.get_array(file_name=qa_data_path)[1:]
  # Sets initial DDOCNAME
  folio_ddocname = ""
  # Iterates OLDPRD line by line
  for old_prd_line in old_prd_content:
    # If different DDOCNAME
    if folio_ddocname != old_prd_line[0]:
      folio_ddocname = old_prd_line[0]
      # Write DDOCNAME
      worksheet.write(row, 0, folio_ddocname)
      # Checks QA file
      for qa_line in qa_content:
        # If same DDOCNAME to OLDPRD
        if folio_ddocname == qa_line[0]:
          # Writes QA content linked with folio
          worksheet.write(row + qa_row, 2, qa_line[1])
          # Increment QA row
          qa_row += 1
      # Checks NEWPRD file
      for new_prd_line in new_prd_content:
        # If same DDOCNAME to OLDPRD
        if folio_ddocname == new_prd_line[0]:
          # Writes NEWPRD content linked with folio
          worksheet.write(row + new_prd_row, 3, new_prd_line[1])
          # Increment NEWPRD row
          new_prd_row += 1
    # Writes OLDPRD content linked with folio
    worksheet.write(row, 1, old_prd_line[1])
    # Increments row
    row += 1
    # Resets QA and NEWPRD rows
    qa_row = new_prd_row = 0
  # Closes workbook file
  workbook.close()

if __name__ == "__main__":
  main()