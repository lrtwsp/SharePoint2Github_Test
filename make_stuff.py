import pandas as pd
import os
import git
import argparse
import os
from git import exc
from oletools.olevba import VBA_Parser
import json
from datetime import datetime, date, time

def extract_vba_to_text(excel_file, output_dir):
    """
    Extracts VBA code from an Excel file and saves it as text files with appropriate extensions.
    Differentiates between modules, classes, and forms.
    Args:
    excel_file (str): The path to the Excel file.
    output_dir (str): The directory where the VBA code will be saved as text.
    """
    vba_parser = VBA_Parser(excel_file)

    if vba_parser.detect_vba_macros():
        for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
            # Determine the correct extension based on the stream_path
            if 'scripts' in stream_path.lower() or 'modules' in stream_path.lower():
                file_extension = '.bas'  # Standard modules
            elif 'classes' in stream_path.lower():
                file_extension = '.cls'  # Class modules
            elif 'forms' in stream_path.lower():
                file_extension = '.frm'  # Forms
            else:
                file_extension = '.vba'  # Default extension

            with open(os.path.join(output_dir, vba_filename + file_extension), 'w') as file:
                file.write(vba_code)

    vba_parser.close()

def convert_excel_to_csvs(excel_file, csv_output_dir):
    """
    Converts each sheet in an Excel file to a separate CSV file.
    Args:
    excel_file (str): The path to the Excel file.
    csv_output_dir (str): The directory where the CSV files will be saved.
    """
    with pd.ExcelFile(excel_file) as xls:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name)
            csv_file = os.path.join(csv_output_dir, f"{sheet_name}.csv")
            df.to_csv(csv_file, index=False)


class CustomJSONEncoder(json.JSONEncoder):
    """ Custom JSON Encoder to handle non-serializable types """
    def default(self, obj):
        if isinstance(obj, (datetime, date, time)):
            return obj.isoformat()
        # You can add more types here if necessary
        return super(CustomJSONEncoder, self).default(obj)

def convert_excel_to_json(excel_file, json_file):
    """
    Converts an Excel file to a JSON file, handling non-serializable types.
    Args:
    excel_file (str): The path to the Excel file.
    json_file (str): The path where the JSON file will be saved.
    """
    with pd.ExcelFile(excel_file) as xls:
        data = {sheet_name: pd.read_excel(xls, sheet_name).applymap(lambda x: x if not isinstance(x, (datetime, date, time)) else x.isoformat()).to_dict(orient='records') 
                for sheet_name in xls.sheet_names}

    with open(json_file, 'w') as file:
        json.dump(data, file, cls=CustomJSONEncoder, indent=4)

def git_operations(repo_dir, commit_message):
    """
    Performs git add, commit, and push operations in the specified repository directory.
    Args:
    repo_dir (str): The path to the local Git repository.
    commit_message (str): The commit message to use.
    """
    repo = git.Repo(repo_dir)
    repo.git.add(A=True)

    try:
        repo.git.commit('-m', commit_message)
    except exc.GitCommandError as e:
        print(f"Git commit failed: {e}")

    try:
        repo.git.push('origin', 'main')
    except exc.GitCommandError as e:
        print(f"Git push failed: {e}")

def main(excel_files, repo_dir, vba=False, csv=False, json=False):
    for file in excel_files:
        csv_output_dir = os.path.join(repo_dir, 'extracted_data', 'csv')
        json_output_dir = os.path.join(repo_dir, 'extracted_data', 'json')
        vba_output_dir = os.path.join(repo_dir, 'extracted_data', 'vba')

        # Create directories if they don't exist
        for dir in [csv_output_dir, json_output_dir, vba_output_dir]:
            if not os.path.exists(dir):
                os.makedirs(dir)

        if vba:
            extract_vba_to_text(file, vba_output_dir)

        if csv:
            convert_excel_to_csvs(file, csv_output_dir)

        if json:
            json_file = os.path.join(json_output_dir, os.path.splitext(os.path.basename(file))[0] + '.json')
            convert_excel_to_json(file, json_file)

    commit_message = 'Updated files: ' + ', '.join([os.path.basename(f) for f in excel_files])
    git_operations(repo_dir, commit_message)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process Excel files for GitHub.')
    parser.add_argument('excel_files', nargs='+', help='Paths to Excel files')
    parser.add_argument('repo_dir', help='Path to the local Git repository')
    parser.add_argument('--vba', action='store_true', help='Extract VBA code from Excel files')
    parser.add_argument('--csv', action='store_true', help='Convert Excel files to CSV')
    parser.add_argument('--json', action='store_true', help='Convert Excel files to JSON')
    args = parser.parse_args()
    main(args.excel_files, args.repo_dir, vba=args.vba, csv=args.csv, json=args.json)