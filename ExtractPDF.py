import csv
import PyPDF2
import re

# Function to extract reference number and tracking number from PDF and save it to a CSV file
def extract_data_to_csv(pdf_file_path, csv_file_path):
    with open(pdf_file_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        extracted_data = []
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            
            # Assuming you have logic to extract reference number and tracking number
            #print(page.extract_text())
            lines = page.extract_text().splitlines()
            for i in lines:
                if "Ref#: " in i:
                    reference_number = i.split(": ")[1]
                match = re.search(r'\b\d{4} \d{4} \d{4} \d{4} \d{4} \d{2}\b', i)
                if match:
                    #print(match.group(0))
                    tracking_number = match.group(0)
                #print(i)
            print("Reference: " + reference_number)
            print("tracking: " + tracking_number)

            #reference_number = "123"  # Replace with actual reference number extraction logic
            #tracking_number = "456"  # Replace with actual tracking number extraction logic
            extracted_data.append([reference_number, tracking_number])
    with open(csv_file_path, 'w', newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        for pair in extracted_data:
            csv_writer.writerow(pair)

# Call the function to extract data from PDF and save it to a CSV file
extract_data_to_csv(r"\\SERVER2\Tech\USPS_Reference\USPS_Label.pdf", r"\\SERVER2\Tech\USPS_Reference\USPS_Excel.csv")
