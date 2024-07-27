import os
from PyPDF2 import PdfWriter, PdfReader

def protect_pdfs(input_folder, output_folder, password):
    for root, dirs, files in os.walk(input_folder):
        for filename in files:
            if filename.endswith('.pdf'):
                input_path = os.path.join(root, filename)
                
                # Create relative path for output
                rel_path = os.path.relpath(root, input_folder)
                output_path = os.path.join(output_folder, rel_path)
                
                # Create output directory if it doesn't exist
                os.makedirs(output_path, exist_ok=True)
                
                # Create a PDF reader object
                reader = PdfReader(input_path)
                writer = PdfWriter()

                # Add all pages to the writer
                for page in reader.pages:
                    writer.add_page(page)

                # Encrypt the new PDF with the password
                writer.encrypt(password)

                # Save the encrypted PDF
                output_file = os.path.join(output_path, f"encrypted_{filename}")
                with open(output_file, "wb") as f:
                    writer.write(f)
                
                print(f"Encrypted: {input_path} -> {output_file}")

# Example usage
input_folder = r"C:\Users\P0037463\Desktop\Unprotected"
output_folder = r"C:\Users\P0037463\Desktop\Protected"
password = "sw|7P7CIoP(8"

protect_pdfs(input_folder, output_folder, password)