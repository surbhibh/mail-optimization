---

## Project Overview: Automated Offer Letter Generation and Email Delivery System

### Introduction
This project automates the generation and distribution of offer letters for HR departments and staffing agencies. The solution reads employee details from an Excel file, generates personalized offer letters using a Word template, converts them to PDFs, and sends them via email using Outlook's SMTP server. Additionally, the script tracks and updates the status of each email sent.

### Key Features
- **Dynamic Offer Letter Generation**: Automatically generate personalized offer letters using a pre-defined Word template.
- **Automated PDF Conversion**: Seamlessly convert generated Word documents to PDFs for consistent presentation.
- **Email Automation**: Automatically email the generated offer letters as attachments using Outlook's SMTP service.
- **Excel Integration**: Track and update the offer letter status ("Sent", "Failed") directly in the source Excel file.
- **Progress and Time Tracking**: Monitor the time taken for each record and the total time for processing all records.
- **Date Handling**: Automatically insert today’s date into each generated offer letter.

### Prerequisites
Before running the script, ensure that the following dependencies are installed:
```bash
pip install openpyxl docxtpl docx2pdf smtplib
```
For Conda users:
```bash
conda install -c conda-forge openpyxl docxtpl
```

### Project Setup
1. **Prepare the Employee Data Excel File**
    - The Excel file (`employee_data.xlsx`) should contain the following columns:
        - `Timestamp`: Record creation timestamp.
        - `Email Address`: The employee's email address.
        - `Agent Name`: The employee's full name.
        - `Agent Mobile Number`: The employee's mobile number.
        - `Agent Whatsapp Number`: The employee's WhatsApp number.
        - `Agent Mail ID`: A secondary email address.
        - `Process`: The role or process assigned to the employee.
        - `Vendor / Consultancy / Freelance HRs Name`: The vendor or consultancy’s name.
        - `Offer Letter Status`: Status of the offer letter ("Sent", "Failed").
        - `Date of Joining`: The date the employee is expected to join.
        - `Payment Status`: The payment status related to the employee (optional).

2. **Prepare the Word Template**
    - The Word template (`OfferLetterTemplate.docx`) should include placeholders for dynamically inserted values:
        - `{{ Agent_Name }}`
        - `{{ Agent_Mobile_Number }}`
        - `{{ Agent_Whatsapp_Number }}`
        - `{{ Agent_Email }}`
        - `{{ Process }}`
        - `{{ Vendor_Name }}`
        - `{{ Date_of_Joining }}`
        - `{{ Current_Date }}`

3. **Email Configuration**
    - Update the script with your Outlook credentials:
    ```python
    EMAIL_USER = 'your-email@outlook.com'
    EMAIL_PASS = 'your-app-password'  # Use an app password if 2FA is enabled
    ```

### Running the Script
1. Ensure all required dependencies are installed and properly configured.
2. Place the Excel file, Word template, and script in the same directory.
3. Run the script using:
```bash
python send_offer_letters.py
```

### Script Execution Flow
1. The script reads employee data from the provided Excel file.
2. For each employee, the script generates a personalized offer letter in Word format, converts it to a PDF, and sends it via email.
3. The script updates the Excel file with the status of each email ("Sent" or "Failed").
4. The progress and processing time for each record are displayed, along with the total time for all records.

### Customization Options
- **Template Customization**: Add more placeholders to the Word template and map them to additional data fields.
- **Email Content**: Customize the email subject, body, and structure to align with your organization’s communication style.
- **Enhanced Logging**: Modify or expand the logging and tracking functionality to capture more detailed metrics.

### Troubleshooting
- **SMTP Errors**: Ensure that SMTP is enabled for your Outlook account. If 2FA is active, use an app password.
- **PDF Conversion Issues**: Confirm that the `docx2pdf` library is installed and functioning correctly.
- **Skipped Records**: Check the "Offer Letter Status" column in the Excel file for "Sent" labels. The script will skip these records to avoid duplicate emails.

### Potential Enhancements
- **Support for Multiple Email Providers**: Extend support for additional email services like Gmail, Yahoo, etc.
- **Advanced Error Logging**: Implement detailed error tracking and logging for better visibility.
- **Scalability Improvements**: Optimize the script to handle bulk operations and larger datasets more efficiently.

### License and Contributions
This project is open-source and licensed under the MIT License. Contributions and suggestions for improvement are always welcome.

---