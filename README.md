# marks_weekly_note

# Summary
This app has been designed to work within the constraints 
presented by a corporate environment. Therefore, it is not 
to be deployed, rather, to be run on a local machine and not 
consume fiscal or hardware resources. Furthermore, security 
constraints, such as governance, 
limit automation possibilities, such as granting the program
access to a company email account where emails can be pulled 
directly.

# Instructions
1. Add environment variable to configuration
2. Copy/Paste email, being sure to include date in upper right,
into variable named 'email' at top of main.py
3. Run
4. Repeat steps 2 and 3 each week with new weekly email

# Environment variables

Name: CSV_PATH  value: full path to desktop

# Constraints

Security - The app shall not be granted access or entitlement
to any email account. Emails will unfortunately need to be
copy and pasted into the program, then ran locally.

Financial - Both cloud and on prem deployments cost money and
hardware space, so to keep these at a minimum, the program is
to be run locally. Additionally, is will not connect to any
database, rather, will use a simple csv file stored on the 
local desktop to support CRUD operations.