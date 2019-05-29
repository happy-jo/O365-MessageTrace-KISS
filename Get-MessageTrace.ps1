# Get-MessageTrace.ps1
# The user conducting the MessageTrace query must be assigned the "Message Tracking" role within Exchange online.
# This is a api query and is not limited to powershell

$creds = Get-Credential 

$MTHeaderParams = @{DataServiceVersion="2.0";MaxDataServiceVersion="2.0";'Accept-Language'="EN-US";Accept="application/json";'Content-Type'="application/json"} 

$messageTrace = (Invoke-WebRequest -Headers $MTHeaderParams -Uri 'https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace' -Method Get -Credential $creds -UseBasicParsing).content | ConvertFrom-Json 

$report = $messageTrace.d.results
$report | ft Index, MessageTraceId, Received, SenderAddress, FromIP, RecipientAddress, ToIP, Subject, Status, Size
