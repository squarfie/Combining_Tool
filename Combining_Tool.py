#!D:\PROGRAMMING\whonet_column_checker\check\python.exe

import pandas as pd
import os
from tabulate import tabulate
import openpyxl

def check_columns(input_folder, input_file1, input_file2, output_folder):
    input1_path = os.path.join(input_folder, input_file1)
    input2_path = os.path.join(input_folder, input_file2)

    # Load input_file1
    try:
        df_input1 = pd.read_excel(input1_path)
        input1_columns = df_input1.columns.tolist()
    except Exception as e:
        print(f"\n❌ Error reading the first input Excel file: {e}")
        return

    # Load input_file2
    try:
        df_input2 = pd.read_excel(input2_path)
        input2_columns = df_input2.columns.tolist()
    except Exception as e:
        print(f"\n❌ Error reading the second input Excel file: {e}")
        return

    # Align columns
    missing_df1_columns = [col for col in input2_columns if col not in input1_columns]
    missing_df2_columns = [col for col in input1_columns if col not in input2_columns]

    for col in missing_df1_columns:
        df_input1[col] = pd.NA
    for col in missing_df2_columns:
        df_input2[col] = pd.NA

    df_input2 = df_input2[df_input1.columns]

    # Validate AccessionNo
    if 'AccessionNo' not in df_input1.columns or 'AccessionNo' not in df_input2.columns:
        print("❌ 'AccessionNo' column must exist in both files.")
        return

    accession1 = df_input1['AccessionNo'].dropna().astype(str)
    accession2 = df_input2['AccessionNo'].dropna().astype(str)

    # Match & Unmatch
    df1_matched = df_input1[df_input1['AccessionNo'].astype(str).isin(accession2)]
    df2_unmatched = df_input2[~df_input2['AccessionNo'].astype(str).isin(accession1)]
    df1_unmatched = df_input1[~df_input1['AccessionNo'].astype(str).isin(accession2)]

    combined_df = pd.concat([df1_matched, df2_unmatched, df1_unmatched], ignore_index=True)

    # Save output
    output_file_name = f"{os.path.splitext(input_file1)[0]}_combined_output.xlsx"
    output_file = os.path.join(output_folder, output_file_name)

    try:
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            combined_df.to_excel(writer, index=False, sheet_name="Combined Data")
            
            # Write summary to another sheet
            summary_data = {
                "Metric": [
                    "Rows in Input 1",
                    "Rows in Input 2",
                    "Matched rows",
                    "Unmatched from Input 1",
                    "Unmatched from Input 2",
                    "Total Combined Rows"
                ],
                "Count": [
                    len(df_input1),
                    len(df_input2),
                    len(df1_matched),
                    len(df1_unmatched),
                    len(df2_unmatched),
                    len(combined_df)
                ]
            }
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, index=False, sheet_name="Summary")
        
        print(f"\n✅ Combined data saved to: {output_file}")
    except Exception as e:
        print(f"\n❌ Error saving the output file: {e}")
        return

    # Print summary
    print("\n📊 Summary:")
    print(f"- Rows in Input 1: {len(df_input1)}")
    print(f"- Rows in Input 2: {len(df_input2)}")
    print(f"- Matched rows: {len(df1_matched)}")
    print(f"- Unmatched from Input 1: {len(df1_unmatched)}")
    print(f"- Unmatched from Input 2: {len(df2_unmatched)}")
    print(f"- Total Combined Rows: {len(combined_df)}")


if __name__ == "__main__":
    print("=== WHONET Dual File Merger (Based on AccessionNo) ===\n")

    # Ask for folder paths only once
    input_folder = input("📁 Enter the full path to the input folder: ").strip()
    output_folder = input("📁 Enter the full path to the output folder: ").strip()

    while True:
        # Ask for input files each time
        input_file1 = input("\n📄 Enter first Excel filename (e.g., file1.xlsx): ").strip()
        input_file2 = input("📄 Enter second Excel filename (e.g., file2.xlsx): ").strip()

        check_columns(input_folder, input_file1, input_file2, output_folder)

        again = input("\n🔁 Would you like to perform another combination? (y/n): ").strip().lower()
        if again not in ['y', 'yes']:
            print("\n👋 Exiting. Have a great day!")
            break
