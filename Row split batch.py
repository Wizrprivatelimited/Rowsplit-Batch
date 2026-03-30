import pandas as pd
import os

input_file = r"C:\Users\taslim.siddiqui\Downloads\Master_FSC_Sheet_ABC&Sarvodaya New 20.03.xlsx"
output_folder = "/Users/romitbenkar/Downloads/split_output"
column_to_split = "Company1"
excel_row_limit = 1_048_000  # Safe margin below Excel’s max of ~1,048,576

os.makedirs(output_folder, exist_ok=True)

df = pd.read_excel(input_file)

output_rows = []
file_counter = 1
current_row_count = 0

def save_current_batch(rows, counter):
    output_df = pd.DataFrame(rows)
    output_file = os.path.join(
        output_folder, f"Split_data_part_{counter}.xlsx"
    )
    output_df.to_excel(output_file, index=False)
    print(f"Saved: {output_file}")

try:
    for index in df.index:
        row = df.loc[index].copy()
        cell_value = row[column_to_split]

        if pd.isna(cell_value) or str(cell_value).strip() == "":
            output_rows.append(row)
            current_row_count += 1
        else:
            split_values = [v.strip() for v in str(cell_value).split(",") if v.strip()]
            for value in split_values:
                new_row = row.copy()
                new_row[column_to_split] = value
                output_rows.append(new_row)
                current_row_count += 1

        # If we approach the Excel row limit, save and start a new file
        if current_row_count >= excel_row_limit:
            save_current_batch(output_rows, file_counter)
            output_rows = []
            file_counter += 1
            current_row_count = 0

    # Save any remaining rows
    if output_rows:
        save_current_batch(output_rows, file_counter)

except Exception as e:
    print(f"An error occurred: {e}")

print("Processing complete.")