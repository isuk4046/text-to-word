# Installation:
* Install the python-docx library (open cmd, then type "pip install python-docx").
* Download the file
* Place it in a folder with the .text files that you want to convert.

# Example usage:
 Convert all .txt files in current directory
* python text_to_docx.py *.txt

 Convert specific files
* python text_to_docx.py file1.txt file2.txt file3.txt

 Convert files and place results in a specific directory
* python text_to_docx.py *.txt -d output_folder

 Convert specific files with custom output names
* python text_to_docx.py input1.txt input2.txt -o output1.docx output2.docx
