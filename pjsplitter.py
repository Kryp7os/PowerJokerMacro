def split_vba_format(base64_command):
    # Define the maximum line length for VBA
    max_line_length = 500
    
    # Split the base64 command into chunks that fit within the maximum line length
    chunks = [base64_command[i:i+max_line_length] for i in range(0, len(base64_command), max_line_length)]
    
    # Join the chunks with VBA line continuation character and "&" for concatenation
    vba_formatted_command = "\"" + "\" & _\n\"".join(chunks) + "\""
    
    return vba_formatted_command

# Get the base64-encoded command from the user
base64_command = input("Enter the base64-encoded command: ")

# Split the command into VBA format
vba_command = split_vba_format(base64_command)

# Print the VBA-formatted command
print("VBA-formatted command:")
print(vba_command)

# Save the VBA-formatted command to a file
with open("split.txt", "w") as file:
    file.write(vba_command)

print("Output saved to split.txt.")
