import re

# List of simplified input strings
input_strings = [
    "nk_viewing_event string",
    "account_number string",
    "linear_slot_audit_ts timestamp"
]

# Function to convert simplified input to full JSON string format
def convert_to_full_format(s):
    name, type_ = s.split()
    return f"{{ mode: 'REQUIRED', name: '{name}', type: {type_.upper()} }}"

# Function to format the string into multi-line JSON with two spaces indentation
def format_string(s):
    # Replace single quotes with double quotes for values
    s = re.sub(r"'", '"', s)
    # Add double quotes around keys
    s = re.sub(r"(\w+):", r'"\1":', s)
    # Add double quotes around unquoted values
    s = re.sub(r'(\s+)(\w+)(\s*)([},])', r'\1"\2"\3\4', s)
    # Format the string into multi-line JSON with two spaces indentation
    s = re.sub(r'{\s*', '{\n  ', s)  # Indent the first line
    s = re.sub(r',\s*', ',\n  ', s)   # Indent following lines
    s = re.sub(r'\s*}', '\n}', s)     # Adjust closing bracket
    return s.strip(',')

# Convert and format each input string
full_format_strings = [convert_to_full_format(s) for s in input_strings]
formatted_strings = [format_string(s) for s in full_format_strings]

# Combine all formatted strings into a single JSON array
json_output = "[\n  " + ",\n\n  ".join(formatted_strings) + "\n]"

# Print the formatted JSON array
print(json_output)
