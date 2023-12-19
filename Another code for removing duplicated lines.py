def remove_duplicate_lines(file_path):
    lines_seen = set()
    output_lines = []

    with open(file_path, 'r') as file:
        for line in file:
            if line not in lines_seen:
                output_lines.append(line)
                lines_seen.add(line)

    with open(file_path, 'w') as file:
        for line in output_lines:
            file.write(line)

# Specify the path to your text file
file_path = 'path_to_text_file'

# Call the function to remove duplicates
remove_duplicate_lines(file_path)

print('Duplicates removed successfully.')
