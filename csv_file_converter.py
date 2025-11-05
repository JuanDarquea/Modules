import pandas as pd
import os
from tkinter import filedialog
from tkinter import Tk

# File to branch?
def convert_csv_to_xlsx(csv_path):
    """
    Convert CSV file to Excel with the same filename
    
    Args:
        csv_path (str): Full path to the CSV file
    """
    try:
        # Read CSV file
        df = pd.read_csv(csv_path)

        # Get the filename without extension
        file_dir = os.path.dirname(csv_path)
        file_name = os.path.splitext(os.path.basename(csv_path))[0]

        # Crate output path with xlsx extension
        output_path = os.path.join(file_dir, f"{file_name}.xlsx")

        # Convert to excel
        df.to_excel(output_path, index=False)

        print("File successfully converted")
        print(f"Input file = {csv_path}")
        print(f"Output file = {output_path}")
        return output_path
    
    except Exception as e:
        print(f"Erro: e")
        return None
    
def select_csv_file():
    """Open file dialog to select a CSV file"""
    root = Tk()
    root.withdraw() # Hide the main window
    root.attributes('-topmost', True) # Bring dialog to front

    # Open file dialog
    file_path = filedialog.askopenfilename(
        title="Select CSV file to convert", 
        filetypes=[
            ("CSV files", "*.csv"), 
            ("All files", "*.*")
        ],
        initialdir=os.path.expanduser("~") # Start in users home directory
    )

    root.destroy() # Close hidden window

    return file_path

def main():
    """Main execution function"""
    print("Please select a CSV file to convert to excel.")

    # Get filepath from dialog
    csv_path = select_csv_file()

    if csv_path:
        print("\nSelected file: {csv_path}")
        convert_csv_to_xlsx(csv_path)
    else:
        print("No file selected. Exiting!")

if __name__ == "__main__":
    main()