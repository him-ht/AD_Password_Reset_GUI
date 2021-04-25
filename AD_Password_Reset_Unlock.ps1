$runuser=whoami
$scriptpath="C:\temp"
$fileName = "AD_password_Reset"

######################### LOG Functions ############################
Function Start-Log {
    
    #Check if file exists and delete if it does
    If((Test-Path -Path $LogFile)){
        Remove-Item -Path $LogFile -Force
    }
 
    #Create file and start logging
    New-Item -Path $LogPath -Name $LogName –ItemType File |Out-Null
 
    Add-Content -Path $LogFile -Value "***************************************************************************************************"
    Add-Content -Path $LogFile -Value "Started Script $($MyInvocation.MyCommand.Name) at [$([DateTime]::Now)]."
    Add-Content -Path $LogFile -Value "***************************************************************************************************"
}

Function Log {
param(
    [string]$In
)
    Add-Content -Path $LogFile -Value $In
	#write-host "$In"
}

Function End-Log {
    Add-Content -Path $LogFile -Value "***************************************************************************************************"
    Add-Content -Path $LogFile -Value "Finished processing at [$([DateTime]::Now)]."
    Add-Content -Path $LogFile -Value "***************************************************************************************************"
}


####### Log File code section #######

$LogName = "$($fileName)_$(get-date -Format "yyyyMMdd-HHmm").log"

if(!(Test-Path "$scriptpath\AD_Password_Reset_Log"))
{
   New-Item -Path "$scriptpath" -Name "AD_Password_Reset_Log" -ItemType Directory -Force 
}
$LogPath = "$scriptpath\AD_Password_Reset_Log"
$LogFile = $LogPath + "\" + $LogName

Start-Log

[void][System.Reflection.Assembly]::LoadWithPartialName( “System.Windows.Forms”)
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName( “Microsoft.VisualBasic”)

#####Define the form size & placement

$form = New-Object System.Windows.Forms.Form
$form.Width = 950;
$form.Height = 400;
$form.FormBorderStyle = 'Fixed3D'
$form.MaximizeBox = $false
$form.Text = "AD Password Reset Tool";
$form.BackColor = "Lightgray"
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;

##############Define input Label

$userLabel = New-Object “System.Windows.Forms.Label”;
$userLabel.Left = 30;
$userLabel.Top = 20;
$userLabel.AutoSize  = $True
$userLabel.Size = '150, 190'
$Font = New-Object System.Drawing.Font("Times New Roman",15)
$userLabel.Font = $Font
$userLabel.Text = "Enter AD User Id";
$form.Controls.Add($userLabel);

##############Define input Label
$emailLabel = New-Object “System.Windows.Forms.Label”;
$emailLabel.Left = 30;
$emailLabel.Top = 60;
$emailLabel.AutoSize  = $True
$emailLabel.Size = '150, 190'
$Font = New-Object System.Drawing.Font("Times New Roman",15)
$emailLabel.Font = $Font
$emailLabel.Text = "Enter Email Id";
$form.Controls.Add($emailLabel);


############### Define text box

$textBoxuser = New-Object System.Windows.Forms.TextBox
$textBoxuser.Left = 500;
$textBoxuser.Top = 20;
$textBoxuser.width = 300;
$textBoxuser.AutoSize = $True
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textBoxuser.Font = $Font
$form.Controls.Add($textBoxuser)


$textBoxemail = New-Object System.Windows.Forms.TextBox
$textBoxemail.Left = 500;
$textBoxemail.Top = 60;
$textBoxemail.width = 300;
$textBoxemail.AutoSize = $True
$Font = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$textBoxemail.Font = $Font
$form.Controls.Add($textBoxemail)


#Add a OK button
$okButton1 = New-Object System.Windows.Forms.Button
$okButton1.Left = 400;
$okButton1.Top = 250;
$okButton1.Width = 150;
$okButton1.Height = 40;
$okButton1.Text = 'OK'
$Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Bold)
$okButton1.Font = $Font
$okButton1.DialogResult=[System.Windows.Forms.DialogResult]::OK
$okButton1.ForeColor = "Black"
$okButton1.BackColor = "PaleGoldenrod"
$form.Controls.Add($okButton1)


#Add a cancel button
$cancelButton1 = New-Object System.Windows.Forms.Button
$cancelButton1.Left = 700;
$cancelButton1.Top = 250;
$cancelButton1.Width = 150;
$cancelButton1.Height = 40;
$cancelButton1.Text = "Cancel"
$Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Bold)
$cancelButton1.Font = $Font
$cancelButton1.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
$cancelButton1.ForeColor = "Black"
$cancelButton1.BackColor = "PaleGoldenrod"
$form.Controls.Add($cancelButton1)
$cancelButton1.add_Click({$form.Close()})

$box = $form.ShowDialog()

$userid=$textBoxuser.Text
$emailid=$textBoxemail.Text

Log "Run user: $runuser"
Log "Entered UserId: $userid"
Log "Entered Email Address: $emailid"

Import-Module ActiveDirectory

try 
{ 
$sam=(Get-ADUser -Filter 'name -eq $userid -or samaccountname -eq $userid' -Properties samaccountname,displayname | Where-Object {$_.Displayname -eq "$userid" -or $_.samaccountname -eq "$userid"}).samaccountname

$var1=(Get-Date).DayOfWeek 
$var2=(Get-date).Day
$password="$var1&?0$var2"
$pass = (ConvertTo-SecureString $password -AsPlainText -Force)

Set-ADAccountPassword $sam -NewPassword $pass -Reset -PassThru | Set-ADUser -ChangePasswordAtLogon $True

$Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($pass)
$pass = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)


#### MAIL PART ####
$fromaddress = "<>" 
$toaddress = "$emailid" 
$Subject = "Password Reset Successful: $userid" 

$body = "Hello User, <br><br>" 
$body += "AD password has been reset.<br>"
$body += "New Password is: $pass <br>"
$body += "<br>"
$body += "<br> <br>Regards <br>"
$body += "Automation Team<br>"                
$smtpserver = ""

$message = new-object System.Net.Mail.MailMessage 
$message.From = $fromaddress 
$message.To.Add($toaddress)
$message.IsBodyHtml = $True 
$message.Subject = $Subject 
$message.body = $body 
$smtp = new-object Net.Mail.SmtpClient($smtpserver) 
$smtp.Send($message)

$lock=(Get-ADUser $sam -Properties Lockedout).Lockedout
if($lock -eq $true)
{    
Unlock-ADAccount -Identity $sam -Confirm:$false
Write-Host "Account has been Unlocked."
Log "Account has been Unlocked."
}

Write-Host "AD Password Reset successful."
Log "AD Password Reset successful."
End-Log
}

catch
{
  Write-Host "Error in AD password reset."
}
