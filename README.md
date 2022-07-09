# Excel VBA mass Emailer

An VBA Solution which allows you to send out emails en masse from a spreadsheet directly in Excel via Outlook. Emails can be customised by individual and can contain attachments.

## The Spreadsheet Template

![Spreadsheet](/Assets/images/Spreadsheet.jpg)

The standard Template contains columsn Email, First Name, Surname, Email Body, Attachment and Subject. The `Email Body` column denotes which HTML file to use as the body of the email. The `Attachment` column denotes which attachments to included as part of the email.

## Customising the template

Shoudl you wish to add or remove columns, be sure to adjust the variable assignments to compensate fro any changes you make

![offsets](/Assets/images/offsets.jpg)

## Precautions

When using Outlook 365, emails are limited to 30 per hour (and 10,000 per day). Should you attempt to send more than this, your account will be locked for 24 hours. In order to prevent emails being sent too rapdily, a 3 second pause occurs between each email. 

![Wait](/Assets/images/Wait.jpg)

Depending of your Operating systems settings, screensavers and sleep mode may prevent emails being sent. To circumvent this, the script using sendkeys to toggle `numlock`. Remove this if it is not necessary for your OS configuration.

![screensaver](/Assets/images/screensaver.jpg)  
