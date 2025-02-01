import os
import comtypes.client
from tkinter import filedialog, Tk


def select_folder(title="Select Folder"):
    """Opens a dialog to select a folder and returns its path."""
    root = Tk()
    root.withdraw()  # Hide the root window
    folder_selected = filedialog.askdirectory(title=title)
    return folder_selected


def convert_doc_to_docx(doc_path, word):
    """Converts .doc (Word 97-2003) to .docx format."""
    doc_path = os.path.normpath(doc_path)  # Normalize path to prevent issues

    if not os.path.exists(doc_path):
        print(f"❌ Error: File does not exist: {doc_path}")
        return None  # Skip conversion if file is missing

    try:
        print(f"📂 Opening file: {doc_path}")  # Debugging step
        doc = word.Documents.Open(doc_path)
        docx_path = doc_path + "x"  # Change ".doc" to ".docx"
        doc.SaveAs(docx_path, FileFormat=16)  # 16 is the format code for .docx
        doc.Close()
        print(f"✅ Converted to .docx: {docx_path}")
        return docx_path
    except Exception as e:
        print(f"❌ Failed to open {doc_path}: {e}")
        return None


def convert_docx_to_pdf(input_folder, output_folder):
    """Converts all .doc and .docx files in the input_folder to PDFs in output_folder."""
    if not os.path.exists(output_folder):
        # Create the output folder if it doesn't exist
        os.makedirs(output_folder)

    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False  # Run Word in the background

    for file in os.listdir(input_folder):
        doc_path = os.path.join(input_folder, file)
        doc_path = os.path.normpath(doc_path)  # Normalize path

        print(f"🔍 Processing: {doc_path}")  # Debugging step

        if not os.path.exists(doc_path):
            print(f"❌ Skipping missing file: {doc_path}")
            continue

        if file.endswith(".doc"):  # Convert .doc to .docx first
            doc_path = convert_doc_to_docx(doc_path, word)
            if doc_path is None:
                continue  # Skip if conversion failed

        if file.endswith(".docx"):  # Convert .docx to PDF
            pdf_path = os.path.join(
                output_folder, file.replace(".docx", ".pdf"))
            pdf_path = os.path.normpath(pdf_path)  # Normalize path

            try:
                print(f"📄 Converting to PDF: {pdf_path}")
                doc = word.Documents.Open(doc_path)
                # 17 is the PDF format code
                doc.SaveAs(pdf_path, FileFormat=17)
                doc.Close()
                print(f"✅ Converted: {file} → {pdf_path}")
            except Exception as e:
                print(f"❌ Failed to convert {doc_path} to PDF: {e}")

    word.Quit()
    print("\n🎉 Conversion Completed!")


# Select input and output folders
input_folder = select_folder("Select the folder containing Word documents")
output_folder = select_folder("Select the folder to save PDFs")

if input_folder and output_folder:
    convert_docx_to_pdf(input_folder, output_folder)
else:
    print("🚫 Operation cancelled.")
