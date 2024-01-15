import os
import pandas as pd

def process_excel_files(input_folder, output_file):
    try:
        # Step 1: Access the "input" folder and read all Excel files
        excel_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx') or f.endswith('.xls') and not f.startswith('~$')]

        if not excel_files:
            print("No Excel files found in the input folder.")
            return

        # Step 2: Combine all Excel files into one file
        combined_df = pd.concat([pd.read_excel(os.path.join(input_folder, file)) for file in excel_files], ignore_index=True)

        # Step 3: Use the first sheet and B column (header: Email)
        email_column_name = 'Email'
        if email_column_name not in combined_df.columns:
            print(f"Error: '{email_column_name}' column not found in the Excel files.")
            return

        # Step 4: Remove blank spaces in B column with entire rows
        combined_df[email_column_name] = combined_df[email_column_name].str.strip()
        combined_df = combined_df.dropna(subset=[email_column_name])

        # Step 5: Remove duplicate values in B column with entire rows
        combined_df = combined_df.drop_duplicates(subset=[email_column_name])

        # Step 6: Save the processed DataFrame to an Excel file
        combined_df.to_excel(output_file, index=False)

        print(f"Processing completed. Results saved to '{output_file}'.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    # Specify the input folder and output file
    input_folder = "input"
    output_file = "output/combined_output2.xlsx"

    # Create the output folder if it doesn't exist
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    # Run the processing function
    process_excel_files(input_folder, output_file)
