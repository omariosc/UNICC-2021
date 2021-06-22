import os
import xlsxwriter

contentSourcePath = "ContentSource\\"
contentTargetPath = "ContentTarget\\"
contentChangesPath = "ContentChanges\\"

# Writes heading for xlsx workbook
def write_headings(worksheet):
  headings = ["ID", "File Change",
  "Source Title", "Target Title", "Change",
  "Source Organization", "Target Organization", "Change",
  "Source Author", "Target Author", "Change",
  "Source Type", "Target Type", "Change",
  "Source Comments", "Target Comments", "Change",
  "Source Profile", "Target Profile", "Change",
  "Source Major Office", "Target Major Office", "Change",
  "Source Content ID", "Target Content ID", "Change",
  "Source REM Number", "Target REM Number", "Change",
  "Source Donor Type", "Target Donor Type", "Change",
  "Source Subsidiary Donor", "Target Subsidiary Donor", "Change",
  "Source Award Number", "Target Award Number", "Change",
  "Source Donor", "Target Donor", "Change",
  "Source Subject", "Target Subject", "Change",
  "Source Remarks", "Target Remarks", "Change"]
  col = 0
  for heading in headings:
    worksheet.write(0, col, heading)
    col += 1

# Formats file content
def format_content(content):
  # Iterates through entire content
  for i in range(len(content)):
    # If the content does not contain a colon
    if ':' not in content[i]:
      # Appends the line to the previous line
      content[i - 1] += " " + content[i]
      # For every line after this line
      for j in range(i, len(content) - 1):
        # Bring the line back one index
        content[j] = content[j + 1]
  # Removes all excess lines
  return content[:15]

# Formats content line
def format_line(line):
  try:
    # Removes heading, colon and space before the content
    line = line[line.index(':') + 2:]
  # If there is no heading the index for ':' reutrns ValueError
  except ValueError:
    pass
  return line

def main():
  # Creates content change directory if doesn't already exist
  if not os.path.exists(contentChangesPath[:-1]):
    os.makedirs(contentChangesPath[:-1])
  # Creates xlsx workbook and worksheet and writes headings
  workbook = xlsxwriter.Workbook(contentChangesPath + "ContentChanges.xlsx")
  worksheet = workbook.add_worksheet()
  write_headings(worksheet)
  # Creates ContentDiffOverview log file
  overview_file = open("ContentDiffOverview.log", "w")
  # Sets initial row and column for worksheet
  row = 1
  col = 2
  # Iterate through content source directory
  for source_name in os.listdir(contentSourcePath[:-1]):
    # Extracts record id
    record_id = source_name.replace("_DataSource.log", "")
    # Retrieves source file content
    source_file = open(contentSourcePath + source_name)
    source_content = source_file.readlines()
    # Creates target file name
    target_name = record_id + "_DataTarget.log"
    # Retrieves target file content
    target_file = open(contentTargetPath + target_name)
    target_content = target_file.readlines()
    # Writes record id to worksheet
    worksheet.write(row, 0, record_id)
    # Used to confirm file change
    file_change = 0
    # Formats file content
    source_content = format_content(source_content)
    target_content = format_content(target_content)
    # Iterates source and target files line by line
    for source_line, target_line in zip(source_content, target_content):
      # Formats source and target lines
      source_line = format_line(source_line)
      target_line = format_line(target_line)
      # Writes source and target content
      worksheet.write(row, col, source_line)
      col += 1
      worksheet.write(row, col, target_line)
      col += 1
      # Matching lines from both files
      if source_line == target_line:
        worksheet.write(row, col, "N")
      # Different lines from both files
      else:
        worksheet.write(row, col, "Y")
        file_change = 1
      col += 1
    # If no differences in files, write to worksheet and overview file and create NoChange file
    if file_change == 0:
      worksheet.write(row, 1, "N")
      change_file = open(contentChangesPath + record_id + "_NoChange.log", "w")
      change_file.close()
    # If difference in files, write to worksheet and overview file and create DataChange file
    else:
      worksheet.write(row, 1, "Y")
      change_file = open(contentChangesPath + record_id + "_DataChange.log", "w")
      change_file.close()
      overview_file.write("DOC:" + record_id + " data changes logged in " + os.getcwd() + "\\" + contentChangesPath + "ContentChanges.xlsx\n")
    # Increments row and resets column
    row += 1
    col = 2
    # Close source and target files
    source_file.close()
    target_file.close()
  # Closes workbook and overview file
  workbook.close()
  overview_file.close()

if __name__ == "__main__":
  main()
