import aspose.email as ae
import os

#def set_aspose_license():  # Commented out
#    """Sets the Aspose.Email license."""
#    try:
#        license = ae.License()
#        license_path = "Aspose.Email.Python.lic"  # Replace with the actual path to your license file
#
#        if os.path.exists(license_path):
#            license.set_license(license_path)
#            print("Aspose.Email license set successfully.")
#        else:
#            print(f"Error: License file not found at {license_path}")
#    except Exception as e:
#        print(f"Error setting Aspose.Email license: {e}")

def extract_html_from_oft(oft_file_path):
    """
    Extracts the HTML body from an OFT file using Aspose.Email.

    Args:
        oft_file_path: The path to the OFT file.

    Returns:
        The HTML body of the OFT file as a string, or None if an error occurs.
    """
    try:
        # Load the OFT file
        msg = ae.MailMessage.load(oft_file_path)

        # Extract the HTML body
        html_body = msg.html_body  # Changed from body_html to html_body

        return html_body

    except Exception as e:
        print(f"Error extracting HTML from OFT: {e}")
        return None

if __name__ == '__main__':
    # Set the Aspose.Email license
    #set_aspose_license()  # Commented out

    # Example usage (replace with a valid OFT file path)
    oft_file = os.path.join(os.path.dirname(__file__), "V2_ST1517567.oft")  # Relative path
    if os.path.exists(oft_file):
        html = extract_html_from_oft(oft_file)
        if html:
            print("HTML Content:\n", html)
        else:
            print("Failed to extract HTML.")
    else:
        print(f"OFT file not found at: {oft_file}")