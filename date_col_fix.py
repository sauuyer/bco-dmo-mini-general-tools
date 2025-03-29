import os
import pandas as pd

def process_excel_file(file_path, file_number, total_files):
    """ Process an Excel file to ensure the 'Date' column is in YYYY-MM-DD format with status updates. """
    try:
        print(f"[{file_number}/{total_files}] Processing: {file_path}")

        # Load the Excel file
        df = pd.read_excel(file_path, engine='openpyxl')

        # Check if the 'Date' column exists
        if 'Date' in df.columns:
            print(f"   → Found 'Date' column. Checking format...")

            # Convert the 'Date' column to datetime format, forcing errors to NaT
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

            # Remove time values and ensure format is YYYY-MM-DD
            df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')

            # Save the updated file
            df.to_excel(file_path, index=False, engine='openpyxl')
            print(f"   ✅ Updated {file_path}")

        else:
            print(f"   ⚠️ Skipping (No 'Date' column): {file_path}")

    except Exception as e:
        print(f"   ❌ Error processing {file_path}: {e}")

def process_directory(directory):
    """ Recursively process all Excel files in a given directory with status updates. """
    file_list = []
    
    # Gather all .xlsx files
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".xlsx"):
                file_list.append(os.path.join(root, file))

    total_files = len(file_list)

    print(f"📂 Found {total_files} .xlsx files in '{directory}'\n")

    # Process each file
    for i, file_path in enumerate(file_list, start=1):
        process_excel_file(file_path, i, total_files)

    print("\n✅ Processing complete!")

# Set the target directory
target_directory = "."  # Change this to your actual directory

# Run the script
process_directory(target_directory)


