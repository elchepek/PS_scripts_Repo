#Enter Job ID here
$JobID = Read-Host "Please enter Job ID: "
$Token = <TOKEN>

#Set Var
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Bearer $Token")
$headers.Add("ask_credential_on_launch", "false")
$TaskName = 'Print%20Script%20Results'

#Get-Results from Ansible ResApi
$URI = "https://<AnsibleURL>/api/v2/jobs/$($JobId.Trim())/job_events/?event__contains=runner_on_ok&task__contains=$TaskName"
$Results = Invoke-RestMethod -Uri $URI -Headers $headers

#Results on command line
$Results.results.event_data.res.msg

#CSS codes
$header = @"
<style>

    h1 {

        font-family: Arial, Helvetica, sans-serif;
        color: #e68a00;
        font-size: 28px;

    }

    h2 {

        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 16px;

    }
    
   table {
		font-size: 12px;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
	} 
	
    td {
		padding: 4px;
		margin: 0px;
		border: 0;
	}
	
    th {
        background: #395870;
        background: linear-gradient(#49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
	}

    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }

    #CreationDate {

        font-family: Arial, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;

    }

    .FailedStatus {

        color: #ff0000;
    }
  
    .PassedStatus {

        color: #008000;
    }

</style>
"@

#convert to HTML 
$html = $Results.results.event_data.res.msg | ConvertTo-Html -Head $header 
$html | Out-File C:\temp\jobReport.html

#Export to CSV
$csv = $Results.results.event_data.res.msg 
$csv | Export-Csv C:\temp\JobReport.csv -NoTypeInformation