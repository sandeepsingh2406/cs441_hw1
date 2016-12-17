This project contains a main class file, myclass.java that connects to the gmail and sheets api and reads and writes the data.
myclass.java first connects to the gmail account using gmail api and gets all unread messages, marking these messages as read. It then gets the dates of these unread messages, and creates a sheet in google sheets using google api, with the name of sheet being the date of the delivery of the email. Thep program writes the time of delivery and subject of the email into the sheet created.
A new sheet is created when a email with a new date is read by the program.

testcase.ps1 is a script to send email to the gmail account which can be accessed using the above program.

Build.gradle has also been included in the proejct which contains all the required dependencies. 

It can be added to the IDE and directly ru using build.gradle and myclass as the main java file. 

