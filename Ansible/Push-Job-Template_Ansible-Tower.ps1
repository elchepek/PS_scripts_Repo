#Enter Template Job ID here
$JobID = Read-Host "Please enter Job template ID: "
$limit = Read-Host "Please enter Server list FQDN in uppercase separated by commas: " #or CSV file
$inv   = Read-Host "Please enter Inventory ID: "
$cred  = Read-Host "Please enter Credential ID: "
$Token = <TOKEN>

#Set Var
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Bearer $Token")
$headers.Add("ask_credential_on_launch", "false")
$headers.Add("oauth2", "/api/o/")
$headers.Add("login_redirect_override", "true")

$myObject = @{          
  ask_limit_on_launch = 'true'
  ask_inventory_on_launch = 'true'
  ask_credential_on_launch = 'true'

  limit = $limit
  inventory = $inv
  credential = $cred
    
   }
    
$JBody = $myObject | ConvertTo-Json

#Get-Results from Ansible ResApi
$URI = "https://<AnsibleURL>/api/v2/job_templates/$JobID/launch/" 
$Results = Invoke-RestMethod -Uri $URI -Headers $headers -Method POST -ContentType "application/json" -UseBasicParsing -Body $JBody

Write-Output "Launching Workflow $($Results.name) with execution Job ID $($Results.job)"