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
    doc_path = os.path.normpath(doc_path)  # Normalize path

    if not os.path.exists(doc_path):
        print(f"❌ Error: File not found: {doc_path}")
        return None  # Skip missing files

    try:
        print(f"📂 Converting .doc to .docx: {doc_path}")
        doc = word.Documents.Open(doc_path)
        docx_path = doc_path + "x"  # Append 'x' to get .docx
        doc.SaveAs(docx_path, FileFormat=16)  # 16 = .docx format
        doc.Close()
        return docx_path
    except Exception as e:
        print(f"❌ Failed to convert {doc_path} to .docx: {e}")
        return None


def convert_to_pdf(input_folder, output_folder):
    """Converts all .doc and .docx files in input_folder to PDFs in output_folder."""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)  # Create output folder if needed

    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False  # Run Word in background

    for file in os.listdir(input_folder):
        doc_path = os.path.join(input_folder, file)
        doc_path = os.path.normpath(doc_path)  # Normalize path

        if not os.path.exists(doc_path):
            print(f"❌ Skipping missing file: {doc_path}")
            continue

        # Convert .doc to .docx first
        if file.endswith(".doc") and not file.endswith(".docx"):
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
                doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF format
                doc.Close()
                print(f"✅ Converted: {file} → {pdf_path}")
            except Exception as e:
                print(f"❌ Failed to convert {doc_path} to PDF: {e}")

    word.Quit()
    print("\n🎉 Conversion Completed!")


# Select input and output folders
input_folder = select_folder("Select folder containing Word documents")
output_folder = select_folder("Select folder to save PDFs")

if input_folder and output_folder:
    convert_to_pdf(input_folder, output_folder)
else:
    print("🚫 Operation cancelled.")
