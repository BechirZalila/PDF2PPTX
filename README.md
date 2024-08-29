# PDF to PPTX Converter

This Python script converts a PDF file into a PowerPoint (PPTX)
presentation where each page of the PDF becomes a slide. The script
provides options to skip specific pages or the first page of the PDF
during the conversion process.

## Requirements

Before running the script, you need to install the required Python
packages. You can install them using the following command:

```bash
pip install -r requirements.txt
```

## Usage

### Command Line Arguments

 - ``pdf_file``: The input PDF file to be converted.
 - ``pptx_file``: The output PPTX file to be created.
 - ``--skip-first``: *Optional*. If provided, the first page of the PDF will
     be skipped.
 - ``--skip``: *Optional*. A comma-separated list of page numbers to skip
    during the conversion (e.g., ``--skip 2,4,5``).

### Examples

Convert PDF to PPTX, including all pages:

```bash
python pdf_to_pptx.py your_file.pdf your_presentation.pptx
```

Convert PDF to PPTX, skipping the first page:

```bash
python pdf_to_pptx.py your_file.pdf your_presentation.pptx --skip-first
```

Convert PDF to PPTX, skipping specific pages (e.g., 2 and 4):

```bash
python pdf_to_pptx.py your_file.pdf your_presentation.pptx --skip 2,4
```

### Notes

The options ``--skip-first`` and ``--skip`` are **mutually exclusive**, meaning
you cannot use them both in the same command.

The generated PPTX file will have each page of the PDF as a slide,
with any specified pages omitted.

### License

This project is licensed under the MIT License. See the LICENSE file
for details.
