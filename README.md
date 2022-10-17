# aircall-gs-cti-dialer-v2

Install a small application inside Google Sheets that will enable the following functionalities directly inside the Google Sheets interface:

- Embedded Aircall CTI Dialer
- Click-to-Dial
- Create Dialer Campaigns

This is possible because Aircall has a broad and well-document public REST API! This version of the application allows you to share one Google Sheet amongst multiple Aircall Users.

But please keep in mind the following:

- Installation only needs to be performed once — after, the agent can use the application as many times as they want (without needing to install again)
- In order for the script to understand which Aircall User is performing the action, one variable need to be modified in Code.gs
  - This variable will keep track of each Google email address and its associated Aircall User ID
- The application will only recognize numbers that are in the International E.164 phone number format (e.g. +46 for Swedish phone numbers)
