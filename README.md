# Markdownify

When making documentation for taking notes or converting from Microsoft Office documents becomes tedious or you need a quick offline solution without using dodgy websites or installing random python libraries, this is a native python script to turn both modern Excel (.xlsx) and modern Word (.docx) files into markdown (.md) files transferring as much of the styling across as can be achieved.

## Current Features

* Convert Excel file to a basic markdown style table (no merged cells)
* Extract the document content from a Word document

## Planned Features

* Convert an entire Word file into a markdown file
    * Subfolder for all of the images with \<img> links
    * Conversion of embedded tables
    * Heading detection and creation
* Merged cell support with HTML-style table support