$EmailTo = "myself@gmail.com"
$EmailFrom = "me@mydomain.com"
$Subject = "Test" 
$Body = "Test Body" 
$SMTPServer = "smtp.gmail.com" 
$filenameAndPath = "C:\abc.txt"
$SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom,$EmailTo,$Subject,$Body)
$attachment = New-Object System.Net.Mail.Attachment($filenameAndPath)
$SMTPMessage.Attachments.Add($attachment)
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587) 
$SMTPClient.EnableSsl = $true 
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential("username", "password"); 
$SMTPClient.Send($SMTPMessage)
$abc = Read-Host -Prompt 'Press any key'
