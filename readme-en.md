I made a script for searching `.xls` and `.xlsx` files with keywords/phrases.

# Requirements
1. Everything from voidtools, https://www.voidtools.com/support/everything/
(https://en.wikipedia.org/wiki/Everything_(software))
2. python

# What does this do
0. **Fast** - uses _parallel processing_, with 4 search keywords and a 1000 excels took only 180s for me
1. creates an csv file that csv be opened in excel with utf-8 names
2. searches all files from basically a csv file files_to_search.efu (efu is from everything save file)
3. searches for amount of strings in [`keywords` variable here](search.py#L107)
4. removes all files that have no matched keywords

its 

# How-to use
1. insure Python and Everything are installed
2. Download the script
3. Open the the script and scroll down until you find `keywords = ['32']` [here](search.py#L107).
4. Add key phrases by enclosing them in single quotes and separating them with commas.
5. Open [Everything tool](https://www.voidtools.com/support/everything/).
6. Perform a search for `.xls`, if you turned on regex then `\.xls`, sort them by path.
  - this will return all types of excel files, 
    - both `.xls` and `.xlsx`
    - autosaves of excel
    - backups of excel-like tools
7. Save as `files_to_search.efu`.
   - You can open it in notepad and delete lines where there definitely is nothing, files are sorted by path and name.
8. Put it in the same folder as the program.
9. Make sure you saved the file with the program and the new keywords.
10. Open the console by typing 'cmd' in the file system search bar (for windows)
11. Run the program through the command in the console `python .\search.py`.
12. Wait for the execution to complete.
13. Open the file `output_results.csv` in excel or any text editor
14. work with the results.



# Example  
put keywords here [here](search.py#L107)
```python
keywords = ['32', 'project', 'other project', 'third project', 'maybe']
```
file `files_to_search.efu` is 1057 entries for me (you cna delete the ones in certain dirs, like trash)
the whole search took 

### cmd output example
Time 23.49 - is total elapsed time.
```
Completed: 307/1057 files (16.54% done). Time 23.49 Estimated time remaining: 118.49 seconds.
Completed: 308/1057 files (16.56% done). Time 24.16 Estimated time remaining: 121.73 seconds.
Completed: 309/1057 files (16.58% done). Time 24.36 Estimated time remaining: 122.56 seconds.
```

```
Completed: 932/1057 files (82.32% done). Time 165.84 Estimated time remaining: 35.61 seconds.
search done
```
This tells us that 1057-932= 125 files had errors for opening, and that the error files are 18% of total size. 

### file output example
```
Filepath;Is_Error;32;project;other project;third project;maybe;Found
A:\$RECYCLE.BIN\S-1-5-21-1075076922-2102090824-4112266334-1001\$R11GD3U.xlsx;False;True;False;False;False;False;1
A:\$RECYCLE.BIN\S-1-5-21-1075076922-2102090824-4112266334-1001\$R4EKPMJ.xlsx;False;True;False;False;False;False;1
A:\$RECYCLE.BIN\S-1-5-21-1075076922-2102090824-4112266334-1001\$R6OFF4A.xlsx;False;True;False;False;False;False;1
```

this shows if there are any errors, the file is sorted by found count.



# Instruction ru
Скрипт для поиска файла с ключевыми словами/фразами

Вот инструкция
1. в майкрософт стор поставить расширение python
2. скачать файл и положить в новую директорию
3. открыть файл и прокрутить вниз пока не найдешь 
  `keywords = ['32'`....
4. добавить ключевые фразы заключив их в одинарные кавычки и разделив запятыми
5. открыть everything
6. сделать поиск .xls, если regex то .xlsx
7. сохранить под названием `files_to_search.efu`
  - можно в блокноте открыть и удалить строки там где точно ничего нет, файлы отсортированы по пути и названию
8. положить в той же папке что и програмку 
9. проверить что сохранили файл с програмкой и новыми ключевыми словами
10. открыть консоль, введя в поисковой строке файловой системы cmd
11. запустить програмку через команду в консоли `python .\search.py`
12. дождаться завершения выполнения
13. открыть в блокноте файл `output_results.csv`, 
14. проверяем по файлам
