
# Excel Files Text Search for Windows and Linux

This script searches `.xls`, `.xlsx`, and `.xlsm` files for specified keywords/phrases.

## Key Features

1. **Fast**: Utilizes parallel processing to expedite searches. With four keywords and 1000 Excel files, it took only 180 seconds.
2. **Configurable**: Specify minimum file size, input filename, output filename, and any number of keywords—all configurable via a file.
3. **Excel-compatible**: Generates a CSV file that can be opened in Excel with UTF-8 characters.
4. **Comprehensive**: Searches all files listed in a CSV file (`files_to_search.efu` from Everything save file).
5. **Efficient**: Excludes files that have no matched keywords from the results.
6. **UTF-8 Support**: Supports filenames and keywords/phrases in UTF-8, allowing for symbols like Cyrillic characters in the search configuration and the resulting CSV output.

## Compatibility

- **Windows**: Designed primarily for Windows.
- **Linux**: Also compatible with Unix-like systems, provided the file list is in the specified format.

## Requirements

1. **Python** with `pandas` (install dependencies via `requirements.txt`).

    ### For Windows
    - [Everything](https://www.voidtools.com/support/everything/) from Voidtools.

    ### For Linux
    - Create a CSV file with filenames and their sizes, with a header like:
      ```
      Filename,Size
      ```

## How to Use

1. Ensure Python and Everything are installed, then run `pip install -r requirements.txt`.
2. Download the script and `search_config.txt`.
3. Open `search_config.txt` and add key phrases after line 19 (after the comments).
4. Use [Everything](https://www.voidtools.com/support/everything/) to search for `.xls` files.
5. Save the search results as `files_to_search.efu`.
6. Place `files_to_search.efu` in the same directory as the script.
7. Ensure `search_config.txt` is correctly configured.
8. Open a console and navigate to the script directory.
9. Run the script with `python .\search.py`.
10. Wait for the script to complete.
11. Open `output_results.csv` in Excel or a text editor to review the results.

## Example

### Configuration (`search_config.txt`)

Supports phrases in other languages via UTF-8. All the necessary instructions are already there: check out [search_config.txt](search_config.txt).

### Command Output Example

```
reading from: files_to_search.efu, exporting to output_csv.csv
files bigger than 5000
with keywords: ['32', 'project', 'other project', 'third project', 'maybe']
removed 174 files that are smaller than 5000
using 32 workers
1/883 files (0.00% done). Time total 0.08 $R59GSHE.xls
2/883 files (0.01% done). Time total 0.11 $RG6HYYA.xls
...
882/883 files (97.10% done). Time total 91.13 Розклад+.xls
883/883 files (100.00% done). Time total 183.74 темп.xlsx
search done
saving
done
Total files with hits 708, files with errors 5
```

### File Output Example

```
sep=;
Filepath;Error_msg;32;project;other project;third project;maybe;Hits
A:\\\\censored_name.xlsx;;True;True;False;False;False;2
C:\\\ncensored_name.xlsx;;True;True;False;False;False;2
A:\$RECYCLE.BIN\...\$R59GSHE.xls;;True;False;False;False;False;1
```

### Errors Example

```
C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\SyncFusion.XlsIO.Base.dll;**File is not a zip file**;False;False;False;False;False;0
```

Good luck with your search!