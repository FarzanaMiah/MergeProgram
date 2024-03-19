import os
import glob
import subprocess
from filesorting import sort_key
from docxcompose.composer import Composer
from docx import Document as Document_compose

def add_page_break(doc):
    doc.add_page_break()

def create_blank_docx(filename):
    # Create a blank Word document
    doc = Document_compose()
    doc.save(filename)
    

def combine_all_docx(filename_master, files_list):
    #try:
    master = Document_compose(filename_master)
    composer = Composer(master)

    sorted_files = sorted(files_list, key=sort_key)

    for file_path in sorted_files:
        if file_path == sorted_files[-1]:
            continue

        doc_temp = Document_compose(file_path)
        print(doc_temp)
        composer.append(doc_temp)

        if file_path != sorted_files[-1]:
            add_page_break(master)
    composer.save("merged_file.docx")
        
 #   except Exception as e:
 #       raise Exception(f'Failed: Base path {e}')
    

def main():
    current_directory = input('Enter the folder path containing your documents: ')

    if os.path.exists(current_directory):
        filename_master = os.path.join(current_directory, "filemaster.docx")

        # Check if "filemaster.docx" exists
        if not os.path.isfile(filename_master):
            create_blank_docx(filename_master)
            print(f"'filemaster.docx' created successfully!")

        files_list = glob.glob(os.path.join(current_directory, '*.docx'))
        print(f"{files_list}")
        combine_all_docx(filename_master, files_list)
        subprocess.run(["open", "merged_file.docx"])
    else:
        raise FileNotFoundError('No such file path or directory')

if __name__ == "__main__":
    main()
