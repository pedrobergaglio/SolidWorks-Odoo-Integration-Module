import re

def find_product_code(error_message):
    # Define the regex pattern to find the code starting with 'W' followed by digits
    pattern = r"W\d+"
    
    # Search for the pattern in the error message
    match = re.search(pattern, error_message)
    
    # If a match is found, return the code
    if match:
        return match.group(0)
    else:
        return None

# Example usage
error_message = "El producto está duplicado. El código de producto al que corresponde el nombre es: W00000116"
code = find_and_save_code(error_message)

if code:
    print(f"Found code: {code}")
    # Save the code to a file or database
    with open("codes.txt", "a") as file:
        file.write(code + "\n")
else:
    print("No code found")
