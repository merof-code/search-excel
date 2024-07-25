import csv
import os
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import ThreadPoolExecutor, as_completed
import time

import warnings

# Suppress specific warnings from openpyxl
warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")
warnings.filterwarnings("ignore", message="Cannot parse header or footer so it will be ignored")

def attempt_read(file_path, engine):
    try:
        xls = pd.ExcelFile(file_path, engine=engine)
        all_sheets_content = []
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine)
            all_sheets_content.append(df.to_string())
        return "\n".join(all_sheets_content), None
    except Exception as e:
        return None, str(e)

def search_in_file(file_path, search_texts):
    # Determine the engine based on file extension
    if file_path.endswith('.xls'):
        primary_engine = 'xlrd'
        secondary_engine = 'openpyxl'
    elif file_path.endswith('.xlsx'):
        primary_engine = 'openpyxl'
        secondary_engine = 'xlrd'
    else:
        primary_engine = 'openpyxl'
        secondary_engine = 'xlrd'
        print(f"***{file_path} is unsupported, will attempt to read anyway ***")
    
    df_string, first_error = attempt_read(file_path, primary_engine)
    if first_error:
        # print(file_path, first_error)
        # If the primary engine fails, try the secondary engine
        df_string, error = attempt_read(file_path, secondary_engine)
        if error:
            # If both attempts fail, return the error
            print(file_path, first_error)
            return file_path, first_error, {text: False for text in search_texts}
    found_texts = {text: text in df_string for text in search_texts}
    return file_path, None, found_texts
    

def get_file_names(csv_file_path):
  with open(csv_file_path, newline='', encoding='utf-8') as csv_file:
          # Create a CSV reader object
          csv_reader = csv.reader(csv_file)

          # Read the header row (optional, remove if there's no header)
          headers = next(csv_reader)
          
          # Check if the first header is 'Filename' or adjust as needed
          if headers[0].lower() != 'filename':
              raise ValueError("First column is not 'Filename'")

          return [{'Filename': row[0], 'Size': int(row[1])} for row in csv_reader]
        
def search_excel_files(csv_file, search_texts, size_threshold = 1000):
    filenames_raw = get_file_names(csv_file)
    filenames = [f for f in filenames_raw if f['Size'] >= size_threshold]
    total_files = len(filenames)
    print(f"removed {len(filenames_raw) - len(filenames)} that are smaller then {size_threshold}")
    total_size = sum(file['Size'] for file in filenames)
    cpu_cores = os.cpu_count()
    # A common heuristic for I/O-bound tasks is to use 2 to 4 times the number of CPU cores
    workers = cpu_cores * 2
    print(f"using {workers} workers")
    with ThreadPoolExecutor(max_workers=workers) as executor:  
        # futures = {executor.submit(search_in_file, file['Filename'], search_texts): file for file in filenames if file['Filename'].endswith(('.xlsx', '.xls'))}
      futures = {executor.submit(search_in_file, file['Filename'], search_texts): file for file in filenames}
      start_time = time.time()
      completed_files = 0
      completed_size = 0

      results = []  # List to store the results
      for future in as_completed(futures):
        result = future.result()  # Get the result from the future
        results.append(result)  # Add the result to the results list

        file = futures[future]
        completed_files += 1
        completed_size += file['Size']
        elapsed_time = time.time() - start_time
        
        percent_done = (completed_size / total_size) * 100
        filename = file['Filename'].split('/')[-1].split('\\')[-1]
        print(f"{completed_files}/{total_files} files ({percent_done:.2f}% done). Time total {elapsed_time:.2f} {filename}")
    return results
  
def is_file_writable(file_path):
    try:
        # Try to open the file in write mode
        with open(file_path, 'w') as file:
            pass  # We don't actually write anything, just test the opening
        return True
    except PermissionError:
        return False
    except Exception as e:
        print(f"An error occurred: {e}")
        return False

def write_results_to_csv(results, search_texts, output_csv, encoding = 'utf-8'):
    for _ in range(5):
        if is_file_writable(output_csv): break;
        if input("File is not writable. Retry? (y/n): ").lower() != 'y': break;
        
    with open(output_csv, mode='w', newline='', encoding=encoding) as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow(["sep=;"])
        # Write the header
        # writer.writerow(["sep=;"])
        header = ['Filepath', 'Error_msg'] + search_texts + ['Found']
        writer.writerow(header)
        
        # Ensure that 'found_texts' is defined before the list comprehension
        filtered_results = []

        for result in results:
            file_path, is_error, found_texts = result  # Extract values from 'result'
            if any(found_texts.values()) or is_error:
                filtered_results.append(result)
        # Sort the filtered results by the count of True values in found_texts in descending order

        with_sum = [(file_path, is_error, found_texts, sum(found_texts.values())) for file_path, is_error, found_texts in filtered_results]
        sorted_results = sorted(with_sum, key=lambda x: x[3], reverse= True)
        # Write the data
        print("saving")
        
        for file_path, is_error, found_texts, sum_value in sorted_results:
            try:
                row = [file_path, is_error] + [found_texts[text] for text in search_texts] + [sum_value]
            except Exception as e:
                print(f"Could not print one row, e")
            writer.writerow(row)
        print("done")
        hits_count = sum(1 for _, is_error, found_texts, _ in sorted_results if not is_error and any(found_texts.values()))
        errors_count = sum(1 for _, is_error, _, _ in sorted_results if is_error)
        total_lines = len(sorted_results)

        print(f"Total files with hits {hits_count}, files with errors {errors_count}, total lines {total_lines}")


def read_config(file_name):
    with open(file_name, 'r') as file:
        lines = file.readlines()
        
    # Remove lines starting with //
    lines = [line for line in lines if not line.strip().startswith("//")]
    threshold_line = lines[0].strip()
    threshold = int(threshold_line.split('=')[1])
    file_list_file_name_line = lines[1].strip()
    file_list_file_name = file_list_file_name_line.split('=')[1]
    output_csv_line = lines[2].strip()
    output_csv= output_csv_line.split('=')[1]
    keywords = [line.strip() for line in lines[3:]]

    return threshold, file_list_file_name, output_csv, keywords

# read input
config_file_name = 'search_config.txt'
file_size_threshold, file_list_file_name, output_csv, keywords = read_config(config_file_name)
output_csv = output_csv + ".csv"

print(f"reading from: {file_list_file_name}, exporting to {output_csv}")
print(f"files bigger then {file_size_threshold}")
print(f"with keywords: {keywords}")

# the program itself
results = search_excel_files(file_list_file_name, keywords, file_size_threshold)
print("search done")
write_results_to_csv(results, keywords, output_csv, 'utf-8-sig')
