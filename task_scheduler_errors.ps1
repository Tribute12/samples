<#

This script outputs Windows events produced by the Task Scheduler that
returned a non-zero code indicating a failure.

Call this script with a date argument in the format mm/dd/YYYY to get events from that date to the current date.
Call this script with no date to get events from the last 24 hours.
Call this script with the argument 'all' to get all matching events in the event log.

This line outputs in list mode: | Format-List *


EventLog XML Query:
$xmlQuery = @'
<QueryList>
  <Query Id="0" Path="Microsoft-Windows-TaskScheduler/Operational">
    <Select Path="Microsoft-Windows-TaskScheduler/Operational">*[System[(EventID=201)]] 
	and *[EventData[(Data[@Name="ResultCode"]!=0)]] 
	and *[EventData[(Data[@Name="ActionName"]="C:\Windows\SYSTEM32\cmd.exe")]]</Select>
  </Query>
</QueryList>
'@

Get-WinEvent -FilterXML $xmlQuery
#>

Param (
    [string]$EventDate = [datetime]::now.AddDays(-1)
)

if ($EventDate -ne [datetime]::now.AddDays(-1) -And $EventDate -ne 'all') {
    try {
        $EventDate = [datetime]::parseexact($EventDate, 'MM/dd/yyyy', $null)
    }
    catch {
        Write-Host 'Please pass date in the format MM/DD/YYYY.'
        Break
    }
}

if ($EventDate -eq 'all') {
    $EventDate = 0
}

Get-WinEvent -FilterHashtable @{
   LogName='Microsoft-Windows-TaskScheduler/Operational';
   ID=201;
   Data='C:\Windows\SYSTEM32\cmd.exe';
   StartTime=$EventDate;
} | Where-Object -Property Message -Match '2147942401' | Format-List *