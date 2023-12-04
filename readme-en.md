I made a script for searching `.xls` and `.xlsx` (and `.xlsm`) files with keywords/phrases.

# Excel files text search for Windows and Linux
This script searches all files old xls and new xlsx, and xlsm content for unlimited amount of strings\phrases\keywords

  ## What is this, why this
  0. **Fast** - uses _parallel processing_, with 4 search keywords and a 1000 excels took only 180s for me
  1. **Configurable** specify min filesize, input filename, output filename, and any number of keywords, **_And all this in a file! :)_**
  2. **Excel-able** creates a csv file, can be opened in excel **with utf-8 chars**
  3. searches all files from basically a csv file files_to_search.efu (efu is from everything save file)
  5. removes all files that have no matched keywords


## Linux
This was built with windows in mind, but works for unix to, if you supply file list in specified format

## Requirements
1. python (pandas (in requirements.txt))

    ## for windows
    Everything from voidtools, https://www.voidtools.com/support/everything/
    (https://en.wikipedia.org/wiki/Everything_(software))

    ## for Linux 
    create a "csv" file by dumping search with filenames and their sizes like so:
    'Filename,Size' - with this exact header


# How-to use
1. insure Python and Everything are installed, do `pip install -r requirements.txt`
2. Download the script and `search_config.txt`
3. Open the the `search_config.txt`
4. Add key phrases after line 19 after the last // (which are comments)
5. Open [Everything tool](https://www.voidtools.com/support/everything/).
6. Perform a search for `.xls`, if you turned on regex then `\.xls`, sort them by path.
  - this will return all types of excel files, 
    - both `.xls`, `.xlsx` and `.xlsm`, and some files with .xls in the middle of filename
    - autosaves of excel
    - backups of excel-like tools
7. Save as `files_to_search.efu`.
   - You can open it in notepad and delete lines where there definitely is nothing, files are sorted by path and name.
8. Put it in the same folder as the program.
9. Make sure you saved `search_config.txt`
10. (for windows) Open the console by typing 'cmd' in the file system search bar 
11. Run the program through the command in the console `python .\search.py`.
12. Wait for the execution to complete.
13. Open the file `output_results.csv` in excel or any text editor
14. work with the results.



# Example  
using `search_config.txt` in repo


file `files_to_search.efu` is 1057 entries for me (you can delete the ones in certain dirs, like trash)
the whole search took 

### cmd output example
Time 23.49 - is total elapsed time.
```
reading from: files_to_search.efu, exporting to output_csv.csv
files bigger then 5000
with keywords: ['32', 'project', 'other project', 'third project', 'maybe']
removed 174 that are smaller then 5000
using 32 workers
1/883 files (0.00% done). Time total 0.08 $R59GSHE.xls
2/883 files (0.01% done). Time total 0.11 $RG6HYYA.xls
...
461/174 files (48.74% done). Time total 31.52 ***.xls
882/883 files (97.10% done). Time total 91.13 Розклад+.xls
883/883 files (100.00% done). Time total 183.74 темп.xlsx
search done
saving
done
Total files with hits 708, files with errors 5, total lines 713
```

### file output example
```
"sep=;"
Filepath;Error_msg;32;project;other project;third project;maybe;Found
A:\\\\censored_name.xlsx;;True;True;False;False;False;2
C:\\\ncensored_name.xlsx;;True;True;False;False;False;2
A:\$RECYCLE.BIN\...\$R59GSHE.xls;;True;False;False;False;False;1
```

Errors look like this 
```
C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\SyncFusion.XlsIO.Base.dll;**File is not a zip file**;False;False;False;False;False;0
```



### Good luck with searching