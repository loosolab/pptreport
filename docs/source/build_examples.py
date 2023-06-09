import os
import glob
import fitz
from natsort import natsorted
from filetree import get_tree_string
import shutil
import yaml

hline = "--------------------\n\n"


def main():

    # Set options
    dpi = 100
    content_dir = "../../examples/content"

    ##################################################
    # Run all examples
    ##################################################

    # Copy content folder to examples
    os.system(f"cp -r {content_dir} examples/")
    thumbs_files = glob.glob("examples/**/Thumbs.db", recursive=True)
    for thumbs_file in thumbs_files:
        os.remove(thumbs_file)  # remove thumbs files from copied content folder

    # zip content folder
    cmd = "cd examples; zip -r content.zip content"
    print(cmd)
    os.system(cmd)

    # Run all examples
    example_files = glob.glob("examples/*.py")
    example_files = natsorted(example_files)  # make sure example 2 is before example 10
    print(f"Found examples: {example_files}")

    for example_file in example_files:
        cmd = f"cd examples; python {os.path.abspath(example_file)}"
        print(cmd)
        os.system(cmd)

    # Semove any existing png files
    png_files = glob.glob("examples/*.png")
    for png_file in png_files:
        os.remove(png_file)

    # Convert all pdfs to individual pngs per page
    pdf_files = glob.glob("examples/*.pdf")
    for pdf_file in pdf_files:

        doc = fitz.open(pdf_file)
        pages = doc.page_count

        for page_num in range(pages):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(dpi=dpi)

            outfile = pdf_file.replace(".pdf", f"_{page_num+1}.png")
            pix.save(outfile)

    example_titles = yaml.safe_load(open("example_titles.yaml", "r"))

    ##################################################
    # Build the rst file
    ##################################################

    rst_file = open("examples/index.rst", "w")

    rst_file.write("Examples\n=========================\n\n")
    rst_file.write("Example data\n---------------------\n\n")

    # Write file tree
    s = get_tree_string("examples/content")
    f = open("examples/tree.txt", "w")
    f.write(s)
    f.close()

    s = ".. literalinclude:: tree.txt"
    rst_file.write(s + "\n\n")

    # Write download
    rst_file.write("Download the content folder:\n")
    rst_file.write(" :download:`content.zip <content.zip>`\n\n")

    for i, example in enumerate(example_files):

        example_name = example.replace(".py", "")
        example_name_base = os.path.basename(example_name)

        # Write example name title
        rst_file.write(hline)
        title = [f"Example {i+1}"]
        if example_name_base in example_titles:
            title.append(example_titles[example_name_base])
        title_str = ": ".join(title) + "\n"
        rst_file.write(title_str)
        rst_file.write("-" * len(title_str) + "\n\n")

        # Write input code
        rst_file.write("Input (script or json):\n")
        rst_file.write("^^^^^^^^^^^^^^^^^^^^^^^\n")

        s = f"""
.. literalinclude:: {example_name_base}.py
    :caption:

.. literalinclude:: {example_name_base}.json
    :caption:
    :language: json
        """
        rst_file.write(s + "\n")

        # Write result
        rst_file.write("Result:\n")
        rst_file.write("^^^^^^^^\n")

        # Write thumbnails for individual slides
        slide_pngs = glob.glob(f"{example_name}_*.png")
        slide_pngs = natsorted(slide_pngs)  # make sure slide 2 is before slide 10
        for png in slide_pngs:

            s = f"""
.. thumbnail:: {os.path.basename(png)}
    :group: {example_name_base}
    :class: framed
    """
            rst_file.write(s)

        # Add option to download pptx / pdf
        rst_file.write("\n\n")
        rst_file.write(f"| :download:`{example_name_base}.pptx <{example_name_base}.pptx>`\n")
        rst_file.write(f"| :download:`{example_name_base}.pdf <{example_name_base}.pdf>`\n\n")

    rst_file.close()

    ##################################################
    # Clean up
    ##################################################

    # Remove copied content folder
    shutil.rmtree("examples/content")
