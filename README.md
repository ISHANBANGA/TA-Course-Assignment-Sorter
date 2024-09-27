# TA Course Assignment Sorter

This project was created for the **Faculty of Engineering and Applied Science (FEAS)** to help sort and merge Teaching Assistant (TA) details with course assignments using two complex Excel sheets. It summarizes which TAs are assigned to which courses, combining data from different files. 

This was developed as a **temporary solution** to organize the data quickly, and moving forward, I plan to develop a more comprehensive **TA Workbook Project** that will allow non-technical users to sort and manage this data through a user-friendly application.

## Project Overview

The program reads data from two Excel files:
1. **Course Summary File** – contains the course information and student details to be matched.
2. **TA Summary File** – contains the names of individuals (TAs) assigned to courses.

The program merges these two data sets, producing a new summary Excel sheet that consolidates the information and shows which TAs are assigned to which courses.

## Features

- **Excel File Manipulation**: Reads and processes data from multiple Excel sheets using ClosedXML.
- **Merge and Sort**: Combines data from TA and course assignment files and matches the TAs to their respective courses.
- **Header Styling**: Applies custom styling to the header row in the Excel output file.
- **Alternate Row Colors**: Implements alternating row colors for better readability of the output file.

## Skills Demonstrated

- Data Processing and Automation using C#
- Excel Manipulation with **ClosedXML**
- Merging and sorting complex datasets
- Applying Excel cell styling programmatically

## Technologies Used

- **C#**
- **ClosedXML** for Excel manipulation
- **Visual Studio** for development

## Installation and Setup

Clone the repository:

Step 1:    git clone https://github.com/ishanbanga/TA-Course-Assignment-Sorter.git

Step 2:    Open the solution in Visual Studio.

Step 3:    Build and run the project.

Step 4:    Place the two Excel files (TA details and course assignments) in the correct directory, and modify any necessary paths in the code.

Step 5:    Run the program to generate the new merged Excel sheet.

## Usage:
This program was designed to help summarize TA assignments across different courses. The output will be a consolidated Excel file that shows which TAs are assigned to which courses.

## Future Enhancements:
- A user-friendly GUI for non-technical users.
- Additional sorting/filtering options.
- Integration with databases for scalability.
