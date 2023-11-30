import pandas as pd
import os
import git
import argparse
from openpyxl import load_workbook
import zipfile
import tempfile
import shutil
import os
from git import Repo, exc
from dotenv import load_dotenv
from oletools.olevba import VBA_Parser

def extract_vba_to_text(excel_file, output_dir):
    """
    Extracts VBA code from an Excel file and saves it as text files.
    Args:
    excel_file (str): The path to the Excel file.
    output_dir (str): The directory where the VBA code will be saved as text.
    """
    vba_parser = VBA_Parser(excel_file)

    if vba_parser.detect_vba_macros():
        for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
            with open(os.path.join(output_dir, vba_filename), 'w') as file:
                file.write(vba_code)

    vba_parser.close()


def extract_vba(excel_file, vba_output_dir):
    """
    Extracts VBA code from an Excel file and saves it to a specified directory.
    Args:
    excel_file (str): The path to the Excel file.
    vba_output_dir (str): The directory where the VBA code will be saved.
    """
    wb = load_workbook(excel_file, keep_vba=True)
    if not wb.vba_archive:
        return

    with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
        wb.save(tmp_file.name)

    vba_folder_name = os.path.splitext(os.path.basename(excel_file))[0] + '_VBA'
    vba_output_path = os.path.join(vba_output_dir, vba_folder_name)

    with zipfile.ZipFile(tmp_file.name, 'r') as zip_ref:
        for file in zip_ref.namelist():
            if file.startswith('xl/vbaProject.bin'):
                zip_ref.extract(file, vba_output_path)

    os.remove(tmp_file.name)

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
        csv_output_dir = os.path.join(repo_dir, 'forgithub')
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
