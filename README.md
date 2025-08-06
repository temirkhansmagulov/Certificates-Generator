# Certificates-Generator
Generate Documents Based on template inside Google Sheets

# Instructions
1) Open: [Certificates Generator Spreadsheet]([url](https://docs.google.com/spreadsheets/d/14MEVZTWE9wdsxmjOI4Y_FqaTaz6sVmGO21nqvd4QFgI/edit?gid=0#gid=0))
2) Make your own copy to Google Drive
3) Edit Templates in the **Settings** Sheet.
   - Basically, you need to make any DOCS type document with variable entries in curly braces {}
   - Choose from this list: {Name and Surname}, {Code},	{Text}, {Signer}, {Value1}, ..., {Value12}
4) Choose where to save files on your Google Drive in the **Settings** Sheet. You have to fill A2 and A4 cells for DOCS and PDF versions respectively.
5) Go to List Sheet and fill out the table with desired data.
6) Select any range (continuous) that contains rows with data you want to use to generate documents in batches
7) Click "Generate Certificates" button at the top.
8) See History Sheet updated with links to your files (PDF versions).
