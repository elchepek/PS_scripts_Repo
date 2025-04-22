#Warning Message
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Pop-up window title
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Auto Assign tool'
$form.Size = New-Object System.Drawing.Size(340,240)
$form.StartPosition = 'CenterScreen'

#Pop-up Botton OK
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

#Pop-up Botton Cancel
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

#Pop-up Label title
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(300,200)
$label.Text = 'Warning: Auto Assign tool is about to start!!!
This tool require that current open Chrome Tab be save of any interaction or script will stop working, you can work in other tabs without any issue. 
Are you ready?'
$form.Controls.Add($label)

#Pop-up List size
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

$Confirm = $form.ShowDialog()

#Options Validation
If ($Confirm -eq 'Ok')

{

#Droplist for Display Name
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Pop-up window title
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Auto Assign tool'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

#Pop-up Botton OK
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

#Pop-up Botton Cancel
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

#Pop-up Label title
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Choose your SNOW display name:'
$form.Controls.Add($label)

#Pop-up List size
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

#Pop-up Options list Assigne
[void] $listBox.Items.Add('John Smith')

#Pop-up selection action
$form.Controls.Add($listBox)
$form.Topmost = $true

#Options Validation
Do {
$Assigne = $null
$result = $null
$result = $form.ShowDialog()
$Assigne = $listBox.SelectedItem

if ($result -eq 'Cancel')
{

Write-Host "Auto assing tool has been cancelled" -ForegroundColor DarkRed
Exit

}

If ($Assigne -eq $null)
{

$ButtonType = [System.Windows.MessageBoxButton]::Ok
$MessageboxTitle = “Auto Assign tool”
$Messageboxbody = “Warning: You must to select one choice"
$MessageIcon = [System.Windows.MessageBoxImage]::Warning
$Confirm1 = [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)

}}

until ($Assigne -ne $null)

#Droplist for time to script to run
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Pop-up window title
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Auto Assign tool'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

#Pop-up Botton OK
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

#Pop-up Botton Cancel
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

#Pop-up Label title
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'At what hour do you want to stop the script?'
$form.Controls.Add($label)

#Pop-up List size
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

#Pop-up Options list time in Hours
[void] $listBox.Items.Add('10')
[void] $listBox.Items.Add('16')
[void] $listBox.Items.Add('19')

#Pop-up selection action
$form.Controls.Add($listBox)
$form.Topmost = $true

#Options Validation
Do {
$LoopHours = $null
$result = $null
$result = $form.ShowDialog()
$LoopHours = $listBox.SelectedItem

if ($result -eq 'Cancel')

{

Write-Host "Auto Assing tool has been cancelled" -ForegroundColor DarkRed
Exit

}

If ($LoopHours -eq $null)
{

$ButtonType = [System.Windows.MessageBoxButton]::Ok
$MessageboxTitle = “Auto Assign tool”
$Messageboxbody = “Warning: You must to select one choice"
$MessageIcon = [System.Windows.MessageBoxImage]::Warning
$Confirm1 = [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)

}}

until ($LoopHours -ne $null)
 
#Filter List Source
$viewurl = 'https://<Instance>.service-now.com/incident_list.do?sysparm_query=active%3Dtrue%5Eassignment_group%3'

#Hide errors when no tickets are pending to assing
$ErrorActionPreference = “SilentlyContinue”
#Open Chrome and Navigate to SNOW
Write-Host "Starting Session" -ForegroundColor Green
#open your profile data, powershell must be running 
$profiledir ="--user-data-dir=$env:USERPROFILE\AppData\Local\Google\Chrome\User Data"
$Chrome = Start-SeChrome #-Arguments ($profiledir)
$Chrome.Navigate().Gotourl('https://<Instance>.service-now.com/sp/?id=sso&portal-id=<Instance>')

#Wait for credentials
Do
{
write-host 'Waiting to reach main console, please Log in to Service Now' -ForegroundColor yellow
Start-Sleep -Seconds 3
}
until ($Chrome.Url.ToString() -match 'https://<Instance>.service-now.com/*') 
Write-Host 'Logged to SNOW' -ForegroundColor Magenta

#Stay alive for selected time (Hours)
$TimeStart = Get-Date
$TimeEnd= Get-Date -Hour $LoopHours -Minute 00 -Second 00
#$TimeEnd = $TimeStart.addhours($LoopHours)
Write-Host "Assigning to $Assigne with role $Userrole"
Write-Host "Start at $TimeStart"
write-host "End at $TimeEnd"

Do { 

#Open selected role view
$Chrome.Navigate().Gotourl($viewurl)
Start-Sleep 2
try {$Chrome.SwitchTo().Frame(0) | Out-Null} catch {}
Write-Host "Assigning  tickets" -ForegroundColor Green

#Checking if the view has any record to assign
#$Tag = 'a'
#$Chrome.FindElementByTagName($Tag) # Use this insty
$XPath = '/html/body/div[1]/div[1]/span/div/div[5]/table/tbody/tr/td/div/table/tbody/tr/td[3]/a'
$found = $Chrome.FindElementbyXpath($XPath)

While ($found)
{

#Grab first ticket on queue
$Incident = $Chrome.FindElementbyXpath($XPath).text
$Chrome.Navigate().Gotourl($Chrome.FindElementbyXpath($XPath).GetAttribute('href'))
Start-Sleep -Seconds 3
$Status = $Chrome.FindElementByName('incident.incident_state').GetAttribute('value')
$Group = $Chrome.FindElementById('sys_display.incident.assignment_group').GetAttribute('value')

#Set ticket in progress
if ($Status -ne '5')
{
#try {$Chrome.SwitchTo().Frame(0) | Out-Null} catch {}
$Chrome.FindElementByName('incident.incident_state').SendKeys("I")
}

#Set Assing to
#try {$Chrome.SwitchTo().Frame(0) | Out-Null} catch {}
$Chrome.FindElementByid('sys_display.incident.assigned_to').SendKeys([OpenQA.Selenium.Keys]::Control+'a')
$Chrome.FindElementById('sys_display.incident.assigned_to').SendKeys($Assigne)
Start-Sleep -Seconds 10
$Chrome.FindElementById('sys_display.incident.assigned_to').SendKeys([OpenQA.Selenium.Keys]::Down)
Start-Sleep -Seconds 2
$Chrome.FindElementById('sys_display.incident.assigned_to').SendKeys([OpenQA.Selenium.Keys]::Enter) 
Start-Sleep -Seconds 2
#comments in ticket
$Chrome.FindElementById('activity-stream-textarea').SendKeys('Ticket has been accepted by script')
$Chrome.FindElementById('activity-stream-textarea').SendKeys(([OpenQA.Selenium.Keys]::Tab)+([OpenQA.Selenium.Keys]::Tab)+([OpenQA.Selenium.Keys]::Enter)) 
#Save ticket
$Chrome.FindElementById('6685b1c93744d7849d3b861754990ef8_bottom').SendKeys([OpenQA.Selenium.Keys]::Enter) 
Start-Sleep -Seconds 3

#Set ticket on hold
$Chrome.FindElementByName('incident.incident_state').SendKeys("O")

#Set on hold reason
$Chrome.FindElementByName('incident.u_on_hold_reasoning').SendKeys("AAA")

#Set on hold date
$holddate = (Get-Date).AddDays(7)  
$Chrome.FindElementByName('incident.u_pending_timer').SendKeys([OpenQA.Selenium.Keys]::Control+'a')
$Chrome.FindElementByName('incident.u_pending_timer').SendKeys($holddate.ToString("yyyy-MM-dd HH:mm:ss" ))
$Chrome.FindElementByName('incident.u_pending_timer').SendKeys([OpenQA.Selenium.Keys]::Enter)

Write-Host "$Incident"

#Add commnets to onhold
$Chrome.FindElementById('activity-stream-textarea').SendKeys('Ticket assigned by script, checking...') 
$Chrome.FindElementById('activity-stream-textarea').SendKeys(([OpenQA.Selenium.Keys]::Tab)+([OpenQA.Selenium.Keys]::Tab)+([OpenQA.Selenium.Keys]::Enter))
$Chrome.FindElementById('sysverb_update_bottom').SendKeys([OpenQA.Selenium.Keys]::Enter)

#Validate if there are ticket remainding
$Chrome.Navigate().Gotourl($viewurl)
$found=$null
try {$Chrome.SwitchTo().Frame(0) | Out-Null} catch {}
$found = $Chrome.FindElementbyXpath($XPath)

}

#If there are no more tickets pending to be assign display message 
$currenttime = get-date
Write-Host "At $currenttime there are no more tickets pending to be assigned, will check in 2 minutes." -ForegroundColor Yellow

#Interval time set in seconds  5 minutes = 300 seconds
Start-Sleep -Seconds 60

$TimeNow = Get-Date
}

#Wait loop
#Until ($TimeNow -ge $TimeEnd)
while ($LoopHours -gt (Get-date -Format HH))

#End
Write-Host "Your shift has ended, get some rest" -ForegroundColor Green
}

Else

{
Write-Host "Auto Assing tool has been cancelled" -ForegroundColor DarkRed
Exit
}