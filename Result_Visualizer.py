import pandas as pd
import matplotlib.pyplot as plt
#Note first run OU_Results_Fetcher.py it will create the excel file we use here
#Crucial changes to make
file_path = r"D:\5th-sem_CSE_Results.xlsx"  # Copy the lacation of the file in your pc and paste in the qoutes
className ="CSE-C 5th Sem" #Enter the name of you class

def Failed_Students_Bar(file_path,className) :
    # Load the data from the Excel file
    data = pd.read_excel(file_path)

    # Print column names to verify them
    print("Column names:", data.columns.tolist())

    # Specify the subject names you want to analyze. Please use the exact same sub names as printed in the memo
    subjects_to_analyze = [
        "OPERATING SYSTEMS", 
        "OPERATING SYSTEM LAB", 
        "COMPUTER ORGANIZATION", 
        "SIGNALS AND SYSTEMS", 
        "MATHEMATICS-III", 
        "COMPUTER ORGANIZATION LAB"
        "DATABASE MANAGEMENT SYS.LAB",
        "EFFECTIVE TECH.COMM.IN ENGLISH",
        "FINANCE AND ACCOUNTING",
        "DATABASE MANAGEMENT SYSTEMS"
    ]

    # Check the actual column names for "Grade Secured" and "Subject Name"
    grade_secured_col = "Grade Secuered"  # Adjust this if the printed column names show a different name
    subject_name_col = "Subject Name"    # Adjust this if the printed column names show a different name

    # Filter the data to include only the rows with the specified subject names
    filtered_data = data[data[subject_name_col].isin(subjects_to_analyze)]

    # Count the number of students who failed (Grade secured is 'F') for each subject name
    failed_counts = filtered_data[filtered_data[grade_secured_col] == 'F'][subject_name_col].value_counts()

    # Create a bar chart
    plt.figure(figsize=(12, 8))
    bars = failed_counts.plot(kind='bar', color='red')
    plt.title("Number of Students Failed in Specified Subjects of "+className)
    plt.xlabel("Subject Name")
    plt.ylabel("Number of Students Failed")
    plt.xticks(rotation=45, ha='right')
    plt.grid(axis='y', linestyle='--', linewidth=0.7)

    # Display the number of students on top of each bar
    for bar in bars.patches:
        bars.annotate(format(bar.get_height(), '.0f'),
                    (bar.get_x() + bar.get_width() / 2, bar.get_height()),
                    ha='center', va='center', xytext=(0, 8),
                    textcoords='offset points')

    plt.tight_layout()
    plt.show()

def Overall_Passed_PieChart(file_path,className):


    # Load the data from the Excel file
    data = pd.read_excel(file_path)

    # Print column names to verify them
    print("Column names:", data.columns.tolist())

    # Ensure that the column names for "Hall Ticket No." and "Final Result" are correct
    hall_ticket_col = "Hall Ticket No."  # Adjust if necessary
    final_result_col = "4th Sem"    # Adjust if necessary

    # Drop duplicate entries based on the Hall Ticket No. to ensure each student is counted only once
    unique_students = data.drop_duplicates(subset=[hall_ticket_col])

    # Count the number of students who are promoted and those who passed
    promoted_count = unique_students[unique_students[final_result_col].str.startswith('PROMOTED--')].shape[0]
    total_students = unique_students.shape[0]
    passed_count = total_students - promoted_count

    # Data for the pie chart
    labels = ['Promoted', 'Passed']
    sizes = [promoted_count, passed_count]
    colors = ['lightcoral', 'lightskyblue']
    explode = (0.1, 0)  # explode the 'Promoted' slice slightly

    # Create a pie chart
    plt.figure(figsize=(8, 8))
    patches, texts, autotexts = plt.pie(sizes, explode=explode, colors=colors, autopct='%1.1f%%',
                                        shadow=True, startangle=140, textprops=dict(color="w"))

    # Add legend with counts
    legend_labels = [f'{label}: {count}' for label, count in zip(labels, sizes)]
    plt.legend(patches, legend_labels, loc="best")

    # Title and formatting
    plt.title('Overall Student Promotion and Passing Status of '+className)
    plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    plt.show()


def process_results(file_path):
    # Load the data from the Excel file to inspect column names
    data = pd.read_excel(file_path)

    # Print column names to verify them
    print("Column names:", data.columns.tolist())

    # Ensure that the column names for "Hall Ticket No." and "Final Result" are correct
    hall_ticket_col = 'Hall Ticket No.'  # Adjust if necessary
    final_result_col = '4th Sem'    # Adjust if necessary

    # Read the Excel file
    df = pd.read_excel(file_path, usecols=[hall_ticket_col, final_result_col], header=None)
    df.columns = ['Roll No.', 'Result']

    # Process roll numbers (last three digits)
    df['Roll No.'] = df['Roll No.'].apply(lambda x: int(x) % 1000)

    # Process results
    def process_result(result):
        if 'PROMOTED--' in result:
            return 0.05
        elif 'PASSED' in result:
            try:
                return float(result.split('-')[-1])
            except ValueError:
                return None
        elif 'PROMOTED-' in result:
            try:
                return float(result.split('-')[-1])
            except ValueError:
                return None
        else:
            return None

    df['SGPA'] = df['Result'].apply(process_result)
    df = df.dropna(subset=['SGPA'])

    # Remove duplicates, keeping the first occurrence
    df = df.drop_duplicates(subset=['Roll No.'], keep='first')

    return df[['Roll No.', 'SGPA']]
def Roll_GPA_Bar(file_path, className):
    # Load the data from the Excel file
    data = pd.read_excel(file_path)

    # Extract the last three digits of the hall ticket number
    data['Roll'] = data['Hall Ticket No.'].astype(str).str[-3:]

    # Ensure the GPA values are floats
    data['GPA'] = data['4th Sem'].apply(lambda x: float(x.split('-')[-1]) if ('PASSED' in x or ('PROMOTED-' in x and not x.endswith('--'))) else (0.5 if 'PROMOTED--' in x else None))

    data = data.dropna(subset=['GPA'])

    # Remove duplicate entries based on the Hall Ticket No.
    data = data.drop_duplicates(subset=['Hall Ticket No.'])

    # Sort the data by Roll
    data = data.sort_values(by='Roll')

    # Create a bar graph of Roll numbers and GPA
    plt.figure(figsize=(12, 8))
    bars = plt.bar(data['Roll'], data['GPA'], color='#f2cbd0')

    # Assign colors based on GPA values
    for bar, gpa in zip(bars, data['GPA']):
        if gpa == 10.0:
            bar.set_color('#FFD700')
        elif 9.0 <= gpa < 10.0:
            bar.set_color('#880d1e')
        elif 8.0 <= gpa < 9.0:
            bar.set_color('#dd2c44')
        elif 7.0 <= gpa < 8.0:
            bar.set_color('#f26a7c')
        elif 6.0 <= gpa < 7.0:
            bar.set_color('#f49ca8')
        elif gpa == 0.5:
            bar.set_color('c')

    # Title and labels
    plt.title(f'Roll Numbers vs GPA of {className} ')
    plt.xlabel('Roll Number')
    plt.ylabel('GPA')
    plt.xticks(rotation=45, ha='right')
    plt.grid(axis='y', linestyle='--', linewidth=0.7)

    plt.tight_layout()
    plt.show()

    # Extract top 3 performers, considering ties
    data = data.sort_values(by='GPA', ascending=False)
    top_performers = data.drop_duplicates(subset=['Hall Ticket No.']).head(5)
    min_gpa = top_performers['GPA'].min()
    additional_top_performers = data[data['GPA'] == min_gpa]
    
    if len(additional_top_performers) > 5:
        top_performers = pd.concat([top_performers, additional_top_performers])
        top_performers = top_performers.drop_duplicates().nlargest(3, 'GPA')

    # Select only the desired columns
    top_performers = top_performers[['Hall Ticket No.', 'Name', 'GPA']]
    print(top_performers)

    # # Save top performers to an Excel file
    # top_performers_file_path = r"E:\top_performers.xlsx"
    # top_performers.to_excel(top_performers_file_path, index=False)

    # print(f"Top 3 performers for {className} have been saved to {top_performers_file_path}")

Roll_GPA_Bar(file_path, className)

# Roll_GPA_Bar(file_path,className)
Overall_Passed_PieChart(file_path,className);
Failed_Students_Bar(file_path,className);
