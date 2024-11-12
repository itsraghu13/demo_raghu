import os
import shutil
from zipfile import ZipFile

# Define source and destination folders
source_folder = '/path/to/network/source_folder'
destination_folder = '/path/to/network/destination_folder'

# Traverse the source folder
for root, dirs, files in os.walk(source_folder):
    for file in files:
        if file.endswith('.archive_extension'):  # replace with your archive file extension
            # Construct full path of the current file
            file_path = os.path.join(root, file)

            # Re-create directory structure in the destination folder
            relative_path = os.path.relpath(root, source_folder)
            destination_dir = os.path.join(destination_folder, relative_path)
            os.makedirs(destination_dir, exist_ok=True)

            # Define path for the compressed file in the destination
            compressed_file_path = os.path.join(destination_dir, f"{file}.zip")

            # Compress the file
            with ZipFile(compressed_file_path, 'w') as zipf:
                zipf.write(file_path, arcname=file)

            print(f"Compressed and moved {file_path} to {compressed_file_path}")
