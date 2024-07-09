# PDF to PPTX Converter

A Python script to convert PDF files to PowerPoint presentations, with each page of the PDF being an image in a slide. The script supports conversion to both JPG and PNG image formats.

## Features

- Converts each page of a PDF file to an image and places it in a PowerPoint slide.
- Supports JPG and PNG image formats.
- Automatically handles file naming and temporary folders.
- Displays progress bars for conversion and slide creation.
- Automatically cleans up temporary files after creating the presentation.
- On macOS, if Microsoft PowerPoint is installed, you can directly input a PPTX file to flatten it.

## Requirements

- Python 3.8 or higher
- Poetry (for dependency management)

## Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/yourusername/pdf_to_pptx.git
    cd pdf_to_pptx
    ```

2. Install Poetry if you haven't already:

    ```bash
    pip install poetry
    ```

3. Install dependencies:

    ```bash
    poetry install
    ```

## Usage

To convert a PDF or PPTX file to a flattened PowerPoint presentation, use the following command:

```bash
poetry run pdf_to_pptx <path/to/your/input_document> --format <image_format>
```

## Arguments

 * `<path/to/your/input_document>`: Path to the input PDF, PPT or PPTX (macOS only) file.
 * `--format <image_format>`: Image format for conversion (jpg or png). Default is jpg.

## Examples

Convert a PDF to a PowerPoint presentation with JPG images:

    ```bash
    poetry run pdf_to_pptx /path/to/your/input_document.pdf --format jpg
    ```

Convert a PDF to a PowerPoint presentation with PNG images:

    ```bash
    poetry run pdf_to_pptx /path/to/your/input_document.pdf --format png
    ```

Convert a PPTX to a PowerPoint presentation (flattened) with JPG images on macOS with Microsoft PowerPoint installed:

    ```bash
    poetry run pdf_to_pptx /path/to/your/input_document.pptx --format jpg
    ```

Convert a PPTX to a PowerPoint presentation (flattened) with PNG images on macOS with Microsoft PowerPoint installed:

    ```bash
    poetry run pdf_to_pptx /path/to/your/input_document.pptx --format png
    ```

> [!NOTE]
> You can use this program to flatten a PowerPoint file by first converting the PowerPoint file to PDF, and then using this program to convert the PDF back to a PowerPoint file. This process makes the PowerPoint content non-editable.
 
## Development

If you want to contribute or modify the script, you can run it in a development environment.

1.	Activate the virtual environment:

    ```bash
    poetry shell
    ```

2.	Run the script:

    ```bash
    python pdf_to_pptx.py <path/to/your/input_document.pdf> --format <image_format>
    ```

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgements

This project uses the following libraries:

- [pdf2image](https://github.com/Belval/pdf2image)
- [python-pptx](https://github.com/scanny/python-pptx)
- [tqdm](https://github.com/tqdm/tqdm)
- [Poetry](https://python-poetry.org/)

## Issues and Contributions

Feel free to submit issues and contribute to the project on [GitHub](https://github.com/inureyes/pdf_to_pptx).
