import sys
import xlsxwriter
import pyexcel as pe

# Writes headings to xlsx workbook
def write_headings(worksheet):
  headings = ["DDOCNAME", "DIFFERENT FIELDS", "OLD PRD DATA", "NEW PRD DATA", "QA DATA"]
  col = 0
  for heading in headings:
    worksheet.write(0, col, heading)
    col += 1

# Formats content line
def format_line(line):
  for i in range(len(line)):
    try:
      # Checks for space from start of string
      while line[i][0] == " ":
        line[i] = line[i][1:]
      # Checks for space from end of string
      while line[i][-1] == " ":
        line[i] = line[i][:-1]
    except (TypeError, IndexError):
      pass
  return line

def main():
  # Gets source and target files from command line arguments
  old_prd_data_path = new_prd_data_path = ""
  try:
    old_prd_data_path = sys.argv[1]
    new_prd_data_path = sys.argv[2]
    # qa_data_path = sys.argv[3]
  except IndexError:
    print("Usage: python physicalitems_diff.py <OLD PRD DATA FILEPATH> <NEW PRD DATA FILEPATH> <QA DATA FILEPATH")
    sys.exit(1)
  # Creates xlsx workbook and worksheet and writes headings
  workbook = xlsxwriter.Workbook((old_prd_data_path.replace("Source", "")).replace(".xlsx", "_Diff.xlsx"))
  worksheet = workbook.add_worksheet("PhysicalDataDiff")
  write_headings(worksheet)
  # Sets initial row and column for worksheet
  row = 1
  # Reads content source and target data from xlsx workbook
  source_content = pe.get_array(file_name=old_prd_data_path)
  target_content = pe.get_array(file_name=new_prd_data_path)
  # Iterates source and target files line by line
  for source_line, target_line in zip(source_content, target_content):
    # Formats source and target lines
    source_line = format_line(source_line)
    target_line = format_line(target_line)
    # If lines are different
    if source_line != target_line:
      # Writes record id to worksheet
      worksheet.write(row, 0, source_line[1])
      # Counter for field index
      field_index = 0
      # Iterates through data cells in line
      for source_cell, target_cell in zip(source_line, target_line):
        # If cells are different
        if source_cell != target_cell:
          # Writes field name to worksheet
          worksheet.write(row, 1, target_content[0][field_index])
          # Writes source and target data to worksheet
          worksheet.write(row, 2, source_cell)
          worksheet.write(row, 3, target_cell)
          row += 1
        # Increments field index
        field_index += 1
  # Closes workbook file
  workbook.close()

if __name__ == "__main__":
  main()