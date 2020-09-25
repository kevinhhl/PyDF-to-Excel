# PyDF to Excel
 A program that converts PDF file into Excel

## Features
- for the backend, it uses **camelot** to read PDF-files and for converting them into DataFrame data structure. It then outputs through Excel via **Pandas**
- the output (Excel workbook) will contain worksheets tabbed out to depict individual pages pertained to the original PDF file
- users can specify the specific pages by entering the page numbers (i.e. “1,2,3-5”)
- simple UI
