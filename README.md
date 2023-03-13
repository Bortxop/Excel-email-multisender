# MULTI SENDER
## Forget about losing time with emails
### Powered by:
[![N|Solid](https://www.versionmuseum.com/images/applications/microsoft-excel/microsoft-excel%5E2016%5Eexcel-logo-new.png)](https://www.microsoft.com/es-es/microsoft-365/excel)

Excel and VBA powered tool to send multiple emails with a custom attached excel file for each receiver. 

- Multiple attached info 
- Progress visualization
- Carefully written comments to understand VBA code

## Requirements

Microsoft Office Excel must be installed. It doesn't work with google sheets or similar.
The functionality has been developed in the 2007 version. There should not be any problem in newer versions.
It hasn't been tested in Excel 2003 and under.

## Usage
### Background
The whole file has been made as an example of usage. In this case, we want to send every week our consumption forecast to various suppliers.

### Workbook sheets

| WORKSHEET NAME | DESCRIPTION |
| ------ | ------ |
| MULTISENDER | Main sheet where data to send is visualised|
| EMAILS | Email per sender|
| DATA | Data from where pivot tables in MULTISENDER take the info|

### Setup

You must provide the data you want to send to each receiver. To do this go to DATA sheet and paste or configure any connection to other excel or access file. After the two lines there are the emails, that are retreived from EMAILS sheet. If the email for each supplier is not found it will raise a cell error statint ("NOT FOUND IN EMAILS")
If you want to change the size of the data columns (the one before the emails, CC and BCC) you would have to change the two pivot table data origins.

Next is to provide the emails in the EMAILS sheet. This can be done manually or with any connection to other files.

Lastly, you have to add the subject, message, attachment name, and the sender (you) server, normally being the sufix after the @ in your email (ie. smtp.outlook.com).
You can optionally add the email and password so it doesn't ask for it everytime you want to send the emails.

And thats all! Update all the pivot tables and workbook sheets with the button provided and hit send.

If done correctly no further setup would be needed. Save the document and don't waste time next time you want to send this repetitive info.

## Troubleshooting
- The smtp port may be different from the default (25). If the error doesn't come from VBA then it may be this issue. 
- Gmail does not allow anymore to provide access to this application (third party programs). If you find anyway to configure this access feel free to provide the info 😉
- If you have little experience with VBA code, it has been decribed, block by block, explaining the secuence followed.

## License

MIT

**Free Tool, Hell Yeah!**
