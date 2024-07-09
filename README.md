# PDF to PPTX Converter

A Python script to convert PDF files to PowerPoint presentations, with each page of the PDF being an image in a slide. The script supports conversion to both JPG and PNG image formats.

## Features

- Converts each page of a PDF file to an image and places it in a PowerPoint slide.
- Supports JPG and PNG image formats.
- Automatically handles file naming and temporary folders.
- Displays progress bars for conversion and slide creation.
- Automatically cleans up temporary files after creating the presentation.

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

To convert a PDF file to a PowerPoint presentation, use the following command:

```bash
poetry run pdf_to_pptx <path/to/your/input_document.pdf> --format <image_format>
```

## Arguments

 * `<path/to/your/input_document.pdf>`: Path to the input PDF file.
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