import pandas as pd
import random
from datetime import datetime, timedelta

# Dummy data generation
num_rows = 20
start_date = datetime(2024, 7, 22, 13, 9, 4)

# Email addresses and agent names to choose from
email_addresses = [
    'chetanrao0486@gmail.com',
    'surbhibh02@gmail.com',
    'surbhibhadvia@gmail.com',
    'junosharma09@gmail.com',
    'surbhibhadviya0912@outlook.com'
]

agent_names = [
    'Surbhi Bh', 'Ruchi Rao', 'Chetan Rao', 'Juno Sharma', 'Neha Verma', 
    'Rajesh Patel', 'Priya Singh', 'Aman Gupta', 'Rohan Desai', 'Pooja Mehta'
]

# Vendors
vendors = ["Ektask Technologies Pvt Ltd", "ABC Solutions", "XYZ Consultancy"]

# Common data
mobile_number = "7999654713"
whatsapp_number = "7999654713"
process = "Flipkart Data Listing & Chat Process"
offer_letter_status = ""
payment_status = ""

# Generating rows
data = []
for i in range(num_rows):
    timestamp = start_date + timedelta(days=random.randint(0, 10), seconds=random.randint(0, 3600))
    email = random.choice(email_addresses)
    agent_name = random.choice(agent_names)
    agent_mail_id = email
    vendor_name = random.choice(vendors)
    date_of_joining = (timestamp + timedelta(days=random.randint(1, 5))).strftime("%m/%d/%Y")
    row = [
        timestamp.strftime("%m/%d/%Y %H:%M:%S"),
        email,
        agent_name,
        mobile_number,
        whatsapp_number,
        agent_mail_id,
        process,
        vendor_name,
        offer_letter_status,
        date_of_joining,
        payment_status
    ]
    data.append(row)

# DataFrame creation
columns = [
    "Timestamp", "Email Address", "Agent Name", "Agent Mobile Number", 
    "Agent Whatsapp Number (Can not be changed)", "Agent Mail ID", 
    "Process", "Vendor / Consultancy / Freelance HRs Name", 
    "Offer Letter Status", "Date of Joining", "Payment Status"
]
df = pd.DataFrame(data, columns=columns)

# Save the DataFrame to an Excel file
df.to_excel("dummy_employee_data_20_rows.xlsx", index=False)
print("Excel file saved successfully!")
