import os
import sys
from PyPDF2 import PdfWriter, PdfReader
from PyPDF2.errors import FileNotDecryptedError, PdfReadError
import msoffcrypto


def encrypt_excel(input_path, output_path, password):
    """Encrypt an Excel file with the given password."""
    # Read the input file
    file = open(input_path, 'rb')
    office_file = msoffcrypto.OfficeFile(file)

    # Save the encrypted file
    with open(output_path, 'wb') as output:
        office_file.encrypt(password, output)

    file.close()


def protect_files(input_folder, output_folder, password):
    skipped_files = []
    processed_files = []

    # Supported file extensions
    supported_extensions = {'.pdf', '.xlsx', '.xls'}

    for root, dirs, files in os.walk(input_folder):
        for filename in files:
            file_ext = os.path.splitext(filename)[1].lower()
            if file_ext in supported_extensions:
                input_path = os.path.join(root, filename)

                try:
                    # Create relative path for output
                    rel_path = os.path.relpath(root, input_folder)
                    output_path = os.path.join(output_folder, rel_path)

                    # Create output directory if it doesn't exist
                    os.makedirs(output_path, exist_ok=True)

                    # Handle based on file type
                    if file_ext == '.pdf':
                        # Process PDF file
                        reader = PdfReader(input_path)

                        # Check if the PDF is encrypted
                        if reader.is_encrypted:
                            print(f"\nFile is already encrypted: {input_path}")
                            try_decrypt = input("Would you like to enter the password to decrypt it? (y/n): ").lower()

                            if try_decrypt == 'y':
                                while True:
                                    pdf_password = input("Enter the PDF password (or 'skip' to skip this file): ")
                                    if pdf_password.lower() == 'skip':
                                        raise ValueError("User chose to skip file")

                                    try:
                                        decrypt_success = reader.decrypt(pdf_password)
                                        if decrypt_success != 0:  # 0 means failed to decrypt
                                            break
                                        print("Incorrect password, please try again")
                                    except:
                                        print("Error decrypting file, please try again")
                            else:
                                raise ValueError("Skipping encrypted file")

                        writer = PdfWriter()
                        for page in reader.pages:
                            writer.add_page(page)

                        writer.encrypt(password)
                        output_file = os.path.join(output_path, f"{filename}")

                        with open(output_file, "wb") as f:
                            writer.write(f)

                    elif file_ext in {'.xlsx', '.xls'}:
                        # Process Excel file
                        output_file = os.path.join(output_path, f"{filename}")
                        encrypt_excel(input_path, output_file, password)

                    processed_files.append(input_path)
                    print(f"Encrypted: {input_path} -> {output_file}")

                except (FileNotDecryptedError, PdfReadError):
                    print(f"Skipping file due to encryption/read error: {input_path}")
                    skipped_files.append(input_path)
                except ValueError as e:
                    print(f"Skipping file: {input_path} - {str(e)}")
                    skipped_files.append(input_path)
                except Exception as e:
                    print(f"Error processing {input_path}: {str(e)}")
                    skipped_files.append(input_path)

    return processed_files, skipped_files


def main():
    if len(sys.argv) != 3:
        print("Usage: python fileprotector.py <input_folder> <password>")
        print("Example: python fileprotector.py \"C:\\Users\\Documents\\Unprotected\" Pass0rd11example")
        print("\nSupported file types: PDF (.pdf), Excel (.xlsx, .xls)")
        sys.exit(1)

    input_folder = sys.argv[1]
    password = sys.argv[2]

    # Get the input folder name and create the output folder name
    folder_name = os.path.basename(input_folder)
    parent_dir = os.path.dirname(input_folder)
    output_folder = os.path.join(parent_dir, f"Protected {folder_name}")

    # Validate input folder exists
    if not os.path.exists(input_folder):
        print(f"Error: Input folder '{input_folder}' does not exist")
        sys.exit(1)

    processed_files, skipped_files = protect_files(input_folder, output_folder, password)

    # Print summary
    print("\nOperation completed!")
    print(f"Successfully processed {len(processed_files)} files")
    print(f"Skipped {len(skipped_files)} files")

    if skipped_files:
        print("\nSkipped files (encrypted, errors, or user-skipped):")
        for file in skipped_files:
            print(f"- {file}")


if __name__ == "__main__":
    main()
