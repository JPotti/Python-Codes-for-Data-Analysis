import re
import pandas as pd

def count_numbers_in_line(line):
    # Count the numbers in the line
    numbers = re.findall(r'\b\d+\.*\d*\b', line)
    return len(numbers)

def extract_fields_from_line(line, patterns):
    # Determine the number of numbers in the line
    num_count = count_numbers_in_line(line)

    # Check if num_count is a valid key in patterns
    if num_count in patterns:
        pattern = patterns[num_count]
        match = re.match(pattern, line)
        if match:
            # Use the appropriate pattern and extract fields
            if num_count == 9:
                data = {
                    'id': int(match.group(1)),
                    'type': int(match.group(2)),
                    'nb': int(match.group(3)),
                    'id_1': int(match.group(4)),
                    'mol': float(match.group(5)),
                    'bo_1': float(match.group(6)),
                    'abo': float(match.group(7)),
                    'nlp': float(match.group(8)),
                    'q': float(match.group(9))
                }
            elif num_count == 11:
                data = {
                    'id': int(match.group(1)),
                    'type': int(match.group(2)),
                    'nb': int(match.group(3)),
                    'id_1': int(match.group(4)),
                    'id_2': int(match.group(5)),
                    'mol': float(match.group(6)),
                    'bo_1': float(match.group(7)),
                    'bo_2': float(match.group(8)),
                    'abo': float(match.group(9)),
                    'nlp': float(match.group(10)),
                    'q': float(match.group(11))
                }
            elif num_count == 13:
                data = {
                    'id': int(match.group(1)),
                    'type': int(match.group(2)),
                    'nb': int(match.group(3)),
                    'id_1': int(match.group(4)),
                    'id_2': int(match.group(5)),
                    'id_3': int(match.group(6)),
                    'mol': float(match.group(7)),
                    'bo_1': float(match.group(8)),
                    'bo_2': float(match.group(9)),
                    'bo_3': float(match.group(10)),
                    'abo': float(match.group(11)),
                    'nlp': float(match.group(12)),
                    'q': float(match.group(13))
                }
            elif num_count == 15:
                data = {
                    'id': int(match.group(1)),
                    'type': int(match.group(2)),
                    'nb': int(match.group(3)),
                    'id_1': int(match.group(4)),
                    'id_2': int(match.group(5)),
                    'id_3': int(match.group(6)),
                    'id_4': int(match.group(7)),
                    'mol': float(match.group(8)),
                    'bo_1': float(match.group(9)),
                    'bo_2': float(match.group(10)),
                    'bo_3': float(match.group(11)),
                    'bo_4': float(match.group(12)),
                    'abo': float(match.group(13)),
                    'nlp': float(match.group(14)),
                    'q': float(match.group(15))
                }
            else:
                return None  # Unknown count

            return data
    else:
        return None

def main(input_file, output_file):
    current_sheet_number = 1
    columns = ['id', 'type', 'nb', 'id_1', 'id_2', 'id_3', 'id_4', 'mol', 'bo_1', 'bo_2', 'bo_3', 'bo_4', 'abo', 'nlp', 'q']

    # Define your patterns
    patterns = {
        9: r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.-]+)',
        11: r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.-]+)',
        13: r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.-]+)',
        15: r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.]+(?:\.\d{1,3})?)\s+([\d.-]+)'
    }

    data = []
    with open(input_file, 'r') as f:
        for line in f:
            line = line.strip()
            if line:
                fields = extract_fields_from_line(line, patterns)
                if fields:
                    data.append(fields)
                if "q" in line:
                    # Create a new sheet when "q" is encountered
                    df = pd.DataFrame(data, columns=columns)
                    with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:
                        sheet_name = f'Sheet_{current_sheet_number}'
                        df.to_excel(writer, index=False, sheet_name=sheet_name)
                    current_sheet_number += 1
                    data = []

        # Create the last sheet if there are remaining data
        if data:
            df = pd.DataFrame(data, columns=columns)
            with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:
                sheet_name = f'Sheet_{current_sheet_number}'
                df.to_excel(writer, index=False, sheet_name=sheet_name)

if __name__ == "__main__":
    input_file = "path_to_input_file"  # Replace with your input file path
    output_file = "path_to_output_file"  # Replace with your desired output file path
    main(input_file, output_file)
