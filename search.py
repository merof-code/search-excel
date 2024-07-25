import csv
import os
import pandas as pd
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
        df_string, error = attempt_read(file_path, secondary_engine)
        if error:
            print(file_path, first_error)
            return file_path, first_error, {text: False for text in search_texts}
    
    found_texts = {text: text in df_string for text in search_texts}
    return file_path, None, found_texts

def get_file_names(csv_file_path):
    with open(csv_file_path, newline='', encoding='utf-8') as csv_file:
        csv_reader = csv.reader(csv_file)
        headers = next(csv_reader)
        if headers[0].lower() != 'filename':
            raise ValueError("First column is not 'Filename'")
        return [{'Filename': row[0], 'Size': int(row[1])} for row in csv_reader]

def search_excel_files(csv_file, search_texts, size_threshold=1000):
    filenames_raw = get_file_names(csv_file)
    filenames = [f for f in filenames_raw if f['Size'] >= size_threshold]
    total_files = len(filenames)
    print(f"Removed {len(filenames_raw) - len(filenames)} files that are smaller than {size_threshold}")
    total_size = sum(file['Size'] for file in filenames)
    cpu_cores = os.cpu_count()
    workers = cpu_cores * 2
    print(f"Using {workers} workers")

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(search_in_file, file['Filename'], search_texts): file for file in filenames}
        start_time = time.time()
        completed_files = 0
        completed_size = 0

        results = []
        for future in as_completed(futures):
            result = future.result()
            results.append(result)
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

def write_results_to_csv(results, search_texts, output_csv, encoding='utf-8'):
    for _ in range(5):
        if is_file_writable(output_csv): break
        if input("File is not writable. Retry? (y/n): ").lower() != 'y': break

    # Prepare data for DataFrame
    data = []
    for result in results:
        file_path, is_error, found_texts = result
        if any(found_texts.values()) or is_error:
            row = [file_path, is_error] + [found_texts[text] for text in search_texts] + [sum(found_texts.values())]
            data.append(row)

    # Create DataFrame
    header = ['Filepath', 'Error_msg'] + search_texts + ['Hits']
    df = pd.DataFrame(data, columns=header)
    df.sort_values(by='Hits', ascending=False, inplace=True)
    # Save to CSV with UTF-8 encoding and BOM
    df.to_csv(output_csv, sep=';', encoding=encoding, index=False)

    hits_count = sum(1 for _, is_error, found_texts in results if not is_error and any(found_texts.values()))
    errors_count = sum(1 for _, is_error, _ in results if is_error)

    print(f"Total files with hits: {hits_count}, files with errors: {errors_count}")


def read_config(file_name):
    with open(file_name, 'r', encoding = 'utf-8') as file:
        lines = file.readlines()
    lines = [line for line in lines if not line.strip().startswith("//")]
    threshold_line = lines[0].strip()
    threshold = int(threshold_line.split('=')[1])
    file_list_file_name_line = lines[1].strip()
    file_list_file_name = file_list_file_name_line.split('=')[1]
    output_csv_line = lines[2].strip()
    output_csv = output_csv_line.split('=')[1]
    keywords = [line.strip() for line in lines[3:]]

    return threshold, file_list_file_name, output_csv, keywords

config_file_name = 'search_config.txt'
file_size_threshold, file_list_file_name, output_csv, keywords = read_config(config_file_name)
output_csv = output_csv + ".csv"

print(f"Reading from: {file_list_file_name}, exporting to {output_csv}")
print(f"Files bigger than {file_size_threshold}")
print(f"With keywords: {keywords}")

results = search_excel_files(file_list_file_name, keywords, file_size_threshold)
print("Search done")
write_results_to_csv(results, keywords, output_csv, 'utf-8-sig')
