![License](https://img.shields.io/github/license/ecommandeur/python-xlsx-conv.svg)

# xlsx-conv

Python script for converting Microsoft Excel Open XML Format Spreadsheets (XLSX/XLSM/XLTX/XLTM files) to Delimiter Separated Values text. Leverages the power of OpenPyXL (http://openpyxl.readthedocs.io/en/default/). 

## Usage

Binaries for Windows and Linux are compiled using PyInstaller. See https://github.com/ecommandeur/python-xlsx-conv/releases .

Run xlsx-conv with -h to get help on commandline options.

```
> xlsx-conv -h
```

Or you can run the script if you are not using the binary.

```
> python xlsx-conv.py -h
```

## Examples

At the minimum, xlsx-conv requires an input.
The input can be a path to an XLSX/XLSM/XLTX or XLTM file.

```
> xlsx-conv -i "resources\MultipleSheets.xlsx"
```

The input can also be a text file specifying paths to one or more XLSX/XLSM/XLTX or XLTM files

```
> xlsx-conv -i "resources\Input.txt"
```

Where Input.txt may looks like

```
Input
temp\Characters.xlsx
temp\Numbers.xlsx
```

If only the input is specified then the output will be written in the same directory as the input.
Also, by default each sheet that is converted will be prefixed by the name of the input. E.g. If MultipleSheets.xlsx contains two sheets named Sheet1 and Sheet2 then by default the ouput will be

```
MultipleSheets.Sheet1.csv
MultipleSheets.Sheet1.csv
```

If you want to have the output in a different directory than you can set the outputDir.

```
> xlsx-conv -i "resources\MultipleSheets.xlsx" -o temp
```

The prefix for the output can be manipulated via `--noprefix` and `--prefix`. The `--noprefix` option will just use the sheet names and the `--prefix` option can be used to set a custom prefix.

```
> xlsx-conv -i "resources\MultipleSheets.xlsx" -o temp --prefix MyPrefix
```

This would output

```
MyPrefix.Sheet1.csv
MyPrefix.Sheet1.csv
```
