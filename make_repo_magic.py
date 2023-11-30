import pandas as pd
import os
import git
import argparse
import os
from git import exc
from oletools.olevba import VBA_Parser

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

def convert_excel_to_csv(excel_file, csv_file):
    """
    Converts an Excel file to a CSV file.
    Args:
    excel_file (str): The path to the Excel file.
    csv_file (str): The path where the CSV file will be saved.
    """
    df = pd.read_excel(excel_file)
    df.to_csv(csv_file, index=False)


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

def main(excel_files, repo_dir):
    for file in excel_files:
        csv_output_dir = os.path.join(repo_dir, 'extracted_data')
        vba_output_dir = csv_output_dir  # Saving VBA code in the same directory

        if not os.path.exists(csv_output_dir):
            os.makedirs(csv_output_dir)

        csv_file = os.path.join(csv_output_dir, os.path.splitext(os.path.basename(file))[0] + '.csv')
        convert_excel_to_csv(file, csv_file)
        extract_vba_to_text(file, vba_output_dir)

    commit_message = 'Updated files: ' + ', '.join([os.path.basename(f) for f in excel_files])
    git_operations(repo_dir, commit_message)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process Excel files for GitHub.')
    parser.add_argument('excel_files', nargs='+', help='Paths to Excel files')
    parser.add_argument('repo_dir', help='Path to the local Git repository')
    args = parser.parse_args()
    main(args.excel_files, args.repo_dir)
