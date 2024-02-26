Email Converter
This is a simple Java application that converts email content to Excel documents.

Features
Converts email content to Excel format
Supports .eml email files
Saves the converted Excel file to the user-specified location
Usage
Open the application: Run the EmailConverter.jar file.
Select the email file: Click the "Browse..." button next to "Open Email File" and choose the .eml email you want to convert.
Select the output directory: Click the "Browse..." button next to "Save Excel File" and choose the directory where you want to save the converted Excel file.
Click "Convert to Excel": The application will process the email and save a new Excel file in the chosen directory, named after the original email with a timestamp.
Requirements
Java 8 or later
Building and Running
Clone the repository:
Bash
git clone https://github.com/<your-username>/email-converter.git
Use code with caution.
Build the project:
Bash
cd email-converter
mvn package
Use code with caution.
This will generate the EmailConverter.jar file in the target directory.
Run the application:
Bash
java -jar target/EmailConverter.jar
Use code with caution.
Contributing
If you would like to contribute to this project, please fork the repository and submit a pull request.

License
This project is licensed under the MIT License.
