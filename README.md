
# Data_Extraction_And_Analysis

## Overview
This project involves analyzing text data extracted from a list of URLs using Python and Microsoft Excel. 
The Python scripts utilize libraries such as requests, BeautifulSoup, NLTK, and pandas for data extraction, text analysis, and data manipulation. 
The extracted data is then stored and analyzed further using Microsoft Excel.

## Key Features
- *Data Extraction:* Python scripts extract text data from URLs and save it as individual text files.
- *Text Analysis:* Sentiment analysis, readability metrics calculation, and additional text statistics computation are performed using Python.
- *Output Generation:* Computed metrics are stored in an Excel file for further analysis and reporting.
- *Error Handling and Validation:* The scripts implement error handling to handle invalid URLs or missing data.
- *Documentation:* Clear documentation of the code is provided for future reference.

## Technologies Used
- *Python:* Utilized for data extraction, text analysis, and script development.
- *Microsoft Excel:* Used for storing and analyzing computed metrics and generating reports.

## Setup
1. *Installation:*
   - Install Python 3.x on your system.
   - Install required Python libraries using pip:
     
     pip install requests beautifulsoup4 nltk pandas
     

2. *Folder Structure:*
   - Organize project files and folders as described in the setup instructions.

3. *Input and Output Files:*
   - Place the input Excel file (Input.xlsx) containing the URLs in the project folder.
   - Prepare the output Excel file (output.xlsx) with the desired structure and place it in the output folder.

4. *Configuration:*
   - Update file paths in Python scripts to point to the correct locations of input and output files and folders.

## Usage
1. *Save Files in a Folder:*
   - Save all the Python files (dataextraction.py, textanalysis.py) in a folder on your computer.

2. *Create Folders for Text Files and Dictionaries:*
   - Create a folder to store all the extracted text files from the URLs.
   - Create a folder named masterdictionary to store the positive and negative word dictionaries.
   - Inside the masterdictionary folder, place the positive and negative word text files.
   - Create a folder named stopwords to store the stopwords text file.

3. *Prepare Output Excel File:*
   - Create an output Excel file (output.xlsx) in the same folder where you have saved the Python files.
   - Ensure that the Excel file has the required structure for storing the relevant output data.

4. *Open Folder in VSCode:*
   - Open Visual Studio Code (VSCode) and navigate to the folder where you have saved all the files.

5. *Install Required Libraries:*
   - Open the terminal in VSCode.
   - Use the following commands to install the required libraries:
     
     pip install requests
     pip install beautifulsoup4
     pip install nltk
     pip install pandas
     
   - These commands will install the necessary libraries (requests, beautifulsoup4, nltk, pandas) for your project.

6. *Replace File Paths:*
   - Open the Python files (dataextraction.py, textanalysis.py) in VSCode.
   - Replace the file paths with the correct paths to the extracted URLs, master dictionary, stopwords, and output Excel file.
   - Ensure that the paths are correctly specified to access the required files.

7. *Run the Program:*
   - After updating the file paths, run the dataextraction.py file first to extract the data from the URLs and save the text files.
   - Then, run the textanalysis.py file to perform text analysis and generate the output data.
   - Make sure to check the output Excel file to verify that the relevant output data has been stored correctly.

## Contributors
- Yashasvi Gupta

## License
This project is licensed under the [MIT License](LICENSE).
