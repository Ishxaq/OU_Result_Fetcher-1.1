import requests
from bs4 import BeautifulSoup
import pandas as pd
#Cucial changes to make!!
url = 'https://www.osmania.ac.in/res07/becbcsaug24.jsp'
start_roll_number = 160422733061  # Replace with the starting roll number
end_roll_number = 160422733120    # Replace with the ending roll number
start_roll_number_LE = 160422733307 # FOR LE students
end_roll_number_LE = 160422733312
RNos = list(range(start_roll_number,end_roll_number+1))

RNos = RNos + list(range(start_roll_number_LE,end_roll_number_LE+1))

def scrape_student_results_between_range(url, RNos):
    # Create an empty DataFrame to store the results
    columns = ["Hall Ticket No.", "Name", "Course", "Subject Code", "Subject Name", "Credits", "Grade Points","Grade Secuered","1st Sem","2nd Sem","3rd Sem"]
    results_data = []
    print(RNos)
    # Iterate through the range of roll numbers
    for roll_number in RNos:
        hall_ticket_number = f"{roll_number:012d}"  # Convert to 12-digit format

        # Send an HTTP request to the URL
        response = requests.post(url, data={'mbstatus': 'SEARCH', 'htno': hall_ticket_number})

        # Check if the request was successful (status code 200)
        if response.status_code == 200:
            # Parse the HTML content of the page using BeautifulSoup
            soup = BeautifulSoup(response.text, 'html.parser')
            #locate the table containing GPA details
            result_details_table = soup.find('table', id='AutoNumber5')
    
            # Debug: Check if the table was found and print its contents
            if result_details_table:
                print(f"AutoNumber5 table found for roll number {hall_ticket_number}")
                # Extract the 2nd and 3rd rows
                rows = result_details_table.find_all('tr')[1:]  # Skip the header row

                result_details = {}
                for row in rows:
                    cols = row.find_all(['td', 'th'])
                    if len(cols) >= 3:
                        key = cols[0].text.strip()
                        value = cols[1].text.strip()
                        result_details[key] = value
                if len(rows) == 2:    
                    key = cols[0].text.strip()
                    value = cols[1].text.strip()
                    result_details[key] = value
                    result_details['1']="N/A"
                    result_details['2']="N/A"



              
            else:
                print(f"AutoNumber5 table not found for roll number {hall_ticket_number}")
                result_details = {}
            # Locate the table containing personal details
            personal_details_table = soup.find('table', id='AutoNumber3')

            # Check if the table was found
            if personal_details_table:
                # Extract data from the personal details table
                rows = personal_details_table.find_all('tr')

                # Dictionary to store personal details
                personal_details = {}

                for row in rows:
                    cols = row.find_all(['td', 'th'])
                    if len(cols) == 4:
                        key = cols[0].text.strip()
                        value = cols[1].text.strip()
                        personal_details[key] = value

                # Locate the table containing marks details
                marks_details_table = soup.find('table', id='AutoNumber4')

                # Check if the table was found
                if marks_details_table:
                    # Extract data from the marks details table
                    rows = marks_details_table.find_all('tr')[1:]  # Skip the header row

                    for row in rows:
                        cols = row.find_all('td')
                        if cols:
                            sub_code = cols[0].text.strip()
                            subject_name = cols[1].text.strip()
                            credits = cols[2].text.strip()
                            grade_secured = cols[4].text.strip()
                            grade_points = cols[3].text.strip()

                            # Append the data to the results list
                            results_data.append([
                                personal_details.get("Hall Ticket No."),
                                personal_details.get("Name"),
                                personal_details.get("Course"),
                                sub_code,
                                subject_name,
                                credits,
                                grade_points,
                                grade_secured,
                                result_details.get("1"),
                                result_details.get("2"),
                                 result_details.get("3")  # Adjust key if necessary
                            ])

        else:
            print(f"Failed to retrieve results for roll number {roll_number}. Status code:", response.status_code)

    # Create a DataFrame from the results list
    results_df = pd.DataFrame(results_data, columns=columns)

    # Save the DataFrame to an Excel file
    results_df.to_excel("REVAL_student_results.xlsx", index=False)
    print("Results saved to student_results.xlsx")

# Example usage
scrape_student_results_between_range(url, RNos)
