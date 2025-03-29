import os
import pandas as pd

def process_excel_file(file_path, file_number, total_files):
    """ Process an Excel file to fix 'Date' and 'Time (UTC-4)' columns, with status updates. """
    try:
        print(f"[{file_number}/{total_files}] Processing: {file_path}")

        # Load the Excel file
        df = pd.read_excel(file_path, engine='openpyxl')

        modified = False  # Track if changes were made

        # ✅ Fix "Date" column (Ensure YYYY-MM-DD format)
        if 'Date' in df.columns:
            print(f"   → Found 'Date' column. Checking format...")
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
            modified = True

        # ✅ Fix "Time (UTC-4)" column (Ensure only time, remove date)
        if 'Time (UTC-4)' in df.columns:
            print(f"   → Found 'Time (UTC-4)' column. Checking format...")
            df['Time (UTC-4)'] = pd.to_datetime(df['Time (UTC-4)'], errors='coerce').dt.strftime('%H:%M:%S')
            modified = True

        # ✅ Save the updated file if any changes were made
        if modified:
            df.to_excel(file_path, index=False, engine='openpyxl')
            print(f"   ✅ Updated {file_path}")
        else:
            print(f"   ⚠️ No changes needed for: {file_path}")

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
