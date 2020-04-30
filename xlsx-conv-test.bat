echo off
REM # 
REM # Batch script for testing xlsx-conv commandline tool on Windows
REM #
REM # We want to test complete batch calls (integral test).
REM # The built-in FC tool can be used to check output against previously converted output that is known to be correct.
REM # FC will set errorlevel to 1 if there are differences
REM # FC will set errorlevel to 2 if it cannot find one or both of the files to compare
REM # dir will set errorlevel to 1 if file is not found (e.g. test for extension or prefix)
REM #
REM # Run by activating xlsx-conv Python environment to load dependencies for xlsx-conv.py
REM # and subsequently calling this script
REM #
if exist temp\ rmdir /S/Q temp\
mkdir temp
REM #
REM # Test simplest call only specifying xlsx/xltx input
REM # -i is the only required argument
REM #
REM # Valuta formatted field in Characters is expected to be dumped without the valuta character!
REM # xlsx-conv dumps the data values
call copy "resources\Characters.xlsx" "temp"
call python xlsx-conv.py -i "temp\Characters.xlsx"
call FC temp\Characters.Characters.csv resources\converted\Characters.Characters.csv
if %ERRORLEVEL%==1 goto failure
if %ERRORLEVEL%==2 goto fcfailure
call copy "resources\Expenses.xltx" "temp"
call python xlsx-conv.py -i "temp\Expenses.xltx"
call FC temp\Expenses.Sheet1.csv resources\converted\Expenses.Sheet1.csv
if %ERRORLEVEL%==1 goto failure
if %ERRORLEVEL%==2 goto fcfailure
REM #
REM # Test different extension
REM #
call python xlsx-conv.py -i "temp\Characters.xlsx" --extension "test.csv"
call dir /B "temp\Characters.Characters.test.csv"
if %ERRORLEVEL%==1 goto failure
REM #
REM # Test different delimiter
REM #
call python xlsx-conv.py -i "temp\Characters.xlsx" --delimiter ";" --extension "semicolon.csv"
call FC temp\Characters.Characters.semicolon.csv resources\converted\Characters.Characters.semicolon.csv
if %ERRORLEVEL%==1 goto failure
if %ERRORLEVEL%==2 goto fcfailure
REM #
REM # Test line break replacement
REM #
call python xlsx-conv.py -i "temp\Characters.xlsx" --linebreak_replacement " " --extension "lr.csv"
call FC temp\Characters.Characters.lr.csv resources\converted\Characters.Characters.lr.csv
if %ERRORLEVEL%==1 goto failure
if %ERRORLEVEL%==2 goto fcfailure
REM #
REM # Test noprefix
REM # Output should be identical to automatically prefixed version
REM #
call python xlsx-conv.py -i "temp\Characters.xlsx" --noprefix
call FC temp\Characters.csv resources\converted\Characters.Characters.csv
if %ERRORLEVEL%==1 goto failure
if %ERRORLEVEL%==2 goto fcfailure
REM #
REM # Test custom prefix
REM # Output should be identical to automatically prefixed version
REM #
call python xlsx-conv.py -i "temp\Characters.xlsx" --prefix "CustomPrefix"
call FC temp\CustomPrefix.Characters.csv resources\converted\Characters.Characters.csv
if %ERRORLEVEL%==1 goto failure
if %ERRORLEVEL%==2 goto fcfailure
REM #
REM # Test output dir
REM #
call python xlsx-conv.py -i "resources\Formulas.xlsx" --o "temp"
call FC temp\Formulas.Formulas.csv resources\converted\Formulas.Formulas.csv
if %ERRORLEVEL%==1 goto failure
if %ERRORLEVEL%==2 goto fcfailure
REM #
REM # Test row and column index
REM #
call python xlsx-conv.py -i "resources\Formulas.xlsx" --o "temp" --col_index --row_index --extension "rc_index.csv"
call FC temp\Formulas.Formulas.rc_index.csv resources\converted\Formulas.Formulas.rc_index.csv
if %ERRORLEVEL%==1 goto failure
if %ERRORLEVEL%==2 goto fcfailure
REM #
REM # Test max columns
REM #
call python xlsx-conv.py -i "resources\Formulas.xlsx" --o "temp" --col_index --row_index --max_cols 3 --extension "max3.csv"
call FC temp\Formulas.Formulas.max3.csv resources\converted\Formulas.Formulas.max3.csv
if %ERRORLEVEL%==1 goto failure
if %ERRORLEVEL%==2 goto fcfailure
REM #
REM # Test TXT input
REM #
REM
REM 
REM
REM # Use ERRORLEVEL to check run
echo %ERRORLEVEL%
if %ERRORLEVEL%==0 goto success

:failure
echo.
echo TEST FAILED 
goto end

:fcfailure
echo.
echo FC could not find converted output in temp or in resources\converted
echo TEST FAILED 
goto end

:success
echo.
echo TEST PASSED 
goto end

:end