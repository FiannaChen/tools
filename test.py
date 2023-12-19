import subprocess
import os
import sys
import glob
import logging
from pathlib import Path
import shutil

def ppt_to_pdf(input_ppt, output_pdf):
    """
    Converts a PowerPoint file to PDF format using unoconv.

    :param input_ppt: Path to the input .ppt or .pptx file.
    :param output_pdf: Path to the output .pdf file.
    """
    try:
        subprocess.run(['unoconv', '-f', 'pdf', '-o', output_pdf, input_ppt], check=True)
        logging.info(f"Conversion successful: {input_ppt} -> {output_pdf}")
    except subprocess.CalledProcessError as e:
        logging.error(f"Error during conversion: {e}")

def main(input_folder, output_folder):
    if shutil.which("unoconv") is None:
        logging.error("unoconv is not installed or not in the PATH.")
        sys.exit(1)

    input_folder = Path(input_folder)
    output_folder = Path(output_folder)

    if not input_folder.exists():
        logging.error(f"Input folder does not exist: {input_folder}")
        sys.exit(1)

    if not output_folder.exists():
        output_folder.mkdir(parents=True)

    files = list(input_folder.glob("*.ppt*"))
    if not files:
        logging.warning("No PowerPoint files found in the input directory.")
        return

    for file in files:
        output_pdf_file = output_folder / (file.stem + ".pdf")
        ppt_to_pdf(str(file), str(output_pdf_file))

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    input_folder = "xxx"
    output_folder = "xxx"
    main(input_folder, output_folder)
