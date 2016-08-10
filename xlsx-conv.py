from openpyxl import load_workbook
import csv
import argparse
import os

# ---
# Get arguments
# ---

parser = argparse.ArgumentParser(description='Convert XLSX file to CSV using openpyxl')
parser.add_argument('-i','--input', help='Full path to XLSX file', required=True)
parser.add_argument('-o','--outputDir', help='Full path output directory', required=False)
parser.add_argument('--noprefix', help='Do not prefix ouput with workbook name', nargs='?', const=1, default=0, required=False)
parser.add_argument('--prefix', help='Use specified prefix instead of prefixing output with workbook name', required=False)
args = parser.parse_args()

inputPath = args.input
outputDir = args.outputDir
noPrefix = args.noprefix
customPrefix = args.prefix
print("noPrefix =", noPrefix)

if not os.path.isfile(inputPath):
    print("Cannot find ", inputPath)
    parser.print_help()
    exit(1)

inputDir, inputFile = os.path.split(inputPath)
inputBaseFn, inputExt = os.path.splitext(inputFile)

if outputDir:
    if not os.path.isdir(outputDir):
        print("Cannot find ", outputDir)
        parser.print_help()
        exit(1)
else:
   outputDir = inputDir

print("Dumping", inputPath)

# ---
# Go ahead and dump that Workbook
# ---

def convertSheet(ws,outputPath):
    with open(outputPath, 'w', encoding='utf-8') as f:
        c = csv.writer(f, lineterminator='\n')
        for row in ws.rows:
            c.writerow([cell.value for cell in row])

wb = load_workbook(filename=inputPath, read_only=True, data_only=True)
ws_names = wb.get_sheet_names()

outPrefix = inputBaseFn + '.'
if customPrefix:
    outPrefix = customPrefix + '.' # override with custom prefix
if noPrefix == 1:
    outPrefix = '' # noprefix takes precedence over customPrefix


for ws_name in ws_names:
    ws = wb[ws_name] # ws is now an IterableWorksheet
    outputPath = outputDir + os.sep + outPrefix + ws_name + '.csv'
    print("Outputting sheet to", outputPath)
    convertSheet(ws,outputPath)

print("Outta here!")