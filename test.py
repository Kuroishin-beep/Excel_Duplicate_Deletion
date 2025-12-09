import pandas as pd
import openpyxl
from tabulate import tabulate
import os

# --- CONFIGURATION: FORCE DISPLAY ---
# This ensures that if you have 500 columns or 5000 rows, Pandas won't hide them.
pd.set_option('display.max_rows', None)  
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 2000) # Wide buffer for terminal

def get_rows_to_delete(df, search_term):
    # 1. Find the main matches (case insensitive search across all columns)
    mask = df.astype(str).apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)
    matched_indices = df[mask].index.tolist()
    
    final_deletion_list = set(matched_indices)
    
    # 2. Look for "Total" rows immediately below matches
    for idx in matched_indices:
        # Check the row immediately below (idx + 1)
        if idx + 1 < len(df):
            next_row = df.iloc[idx + 1]
            row_content = str(next_row.values).lower()
            
            # Logic: If the row below contains "total", add it to kill list
            if "total" in row_content:
                final_deletion_list.add(idx + 1)

    return sorted(list(final_deletion_list))

def delete_rows_preserve_formatting(file_path, indices_to_delete):
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active 
        
        # Convert Pandas Index (0-based) to Excel Row (1-based + Header)
        # We add 2 because Excel Row 1 is header, Row 2 is Index 0.
        excel_rows_to_delete = [i + 2 for i in indices_to_delete]
        
        # IMPORTANT: Sort descending to delete from bottom up
        excel_rows_to_delete.sort(reverse=True)
        
        print(f"\nProcessing deletion on Excel Rows: {excel_rows_to_delete}...")
        
        for r in excel_rows_to_delete:
            ws.delete_rows(r)
            
        output_file = "cleaned_formatted_output.xlsx"
        wb.save(output_file)
        return output_file
        
    except Exception as e:
        print(f"Error saving formatted file: {e}")
        return None

def main():
    # 1. INPUT (Drag and Drop Simulation)
    file_path = input("Enter Excel file path: ").strip().strip('"')
    if not os.path.exists(file_path):
        print("File not found.")
        return

    # Load the data
    df = pd.read_excel(file_path)

    # 2. LEFT MODAL: VIEW ALL DATA (SCROLLABLE)
    print(f"\n--- LEFT MODAL: Current Data ({len(df)} rows) ---")
    print("Listing all rows... (Scroll up in your terminal to view)")
    
    # We add 'showindex=True' so you see the ID number on the far left
    print(tabulate(df.head(50), headers='keys', tablefmt='grid', showindex=True))
    
    print(tabulate(df.iloc[100:150], headers='keys', tablefmt='grid', showindex=True))
    
    print("--- End of Current Data View ---\n")

    # 3. SEARCH FUNCTION
    search_term = input("Enter keyword to delete (e.g., 'INV-2023'): ")
    
    delete_indices = get_rows_to_delete(df, search_term)
    
    if not delete_indices:
        print("No rows found matching that term.")
        return

    # Create a view of what will be deleted
    to_delete_df = df.iloc[delete_indices].copy()
    
    # Add a visual column for "Excel Row Number" to make it easier for the user
    to_delete_df.insert(0, "Excel_Row", [i + 2 for i in delete_indices])


    # 4. MIDDLE MODAL: REVIEW DELETION (SCROLLABLE)
    print(f"\n--- MIDDLE MODAL: Found {len(delete_indices)} rows to delete ---")
    print("These rows (and associated Totals) will be removed:")
    
    print(tabulate(to_delete_df, headers='keys', tablefmt='grid'))
    
    # 5. RESCUE LOGIC (Validation)
    
    print("\nDo you want to REMOVE any rows from this deletion list? (Keep them in the file)")
    print("Enter the 'Excel_Row' number shown above (e.g., 12, 15). Press Enter to skip.")
    
    rescue_input = input("Rescue Rows: ")
    
    final_indices = delete_indices.copy()
    
    if rescue_input:
        try:
            rescue_rows = [int(x.strip()) for x in rescue_input.split(',')]
            # Convert Excel Row back to Pandas Index to remove from logic list
            rescue_indices = [r - 2 for r in rescue_rows]
            final_indices = [i for i in final_indices if i not in rescue_indices]
            print(f"Rescued Excel Rows: {rescue_rows}")
        except ValueError:
            print("Invalid input. Proceeding with full list.")
            
    # 6. EXECUTE AND DOWNLOAD
    if not final_indices:
        print("No rows selected for deletion.")
        return

    confirm = input(f"\nFinal Check: Delete {len(final_indices)} rows? (yes/no): ").lower()
    if confirm == 'yes':
        saved_file = delete_rows_preserve_formatting(file_path, final_indices)
        if saved_file:
            print(f"\nDONE! Download your file here: {saved_file}")
            print("Note: All formatting, colors, and borders are preserved.")
    else:
        print("Operation cancelled.")

if __name__ == "__main__":
    main()