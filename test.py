import openpyxl

def analyze_excel_markers(workbook):
    """Analyze the Excel workbook to get all markers (sheet name, start, and end) and create a mapping."""
    mapping = {}
    try:
        for sheet in workbook.worksheets:
            start_row = None
            end_row = None
            sheet_name = sheet.title

            for row_idx, row in enumerate(sheet.iter_rows(min_col=1, max_col=1, values_only=True), start=1):
                cell_value = row[0]
                if isinstance(cell_value, str) and cell_value.startswith("pptstart:"):
                    marker_name = cell_value.split(":")[1]
                    start_row = row_idx
                elif isinstance(cell_value, str) and cell_value.startswith("pptend:"):
                    end_row = row_idx 

                if start_row and end_row:
                    mapping[marker_name] = {
                        "sheet_name": sheet_name,
                        "start_row": start_row,
                        "end_row": end_row
                    }
                    start_row = None
                    end_row = None

    except Exception as e:
        print(f"Failed to analyze Excel workbook: {e}")

    return mapping

def main():
    # Path to your Excel file
    excel_file_path = '/Users/nourajneid/Downloads/pythonAuthomatisation/MAF Enova WS August 2024-for test.xlsx'

    try:
        # Load the workbook
        workbook = openpyxl.load_workbook(excel_file_path)

        # Analyze the workbook
        mapping = analyze_excel_markers(workbook)

        # Print the mapping
        for marker, details in mapping.items():
            print(f"Marker: {marker}")
            print(f"  Sheet Name: {details['sheet_name']}")
            print(f"  Start Row: {details['start_row']}")
            print(f"  End Row: {details['end_row']}")
            print()

    except FileNotFoundError:
        print(f"The file {excel_file_path} was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
