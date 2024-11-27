$ConfigFile = "$PSScriptRoot\O365Monitor-Config.xml"
$xml = New-Object System.Xml.XmlDocument
$xml.Load($ConfigFile)
$ConfigParams = $xml.SelectSingleNode("//o365app")

$LogFile = "$PSScriptRoot\O365Montior-PublicGroups.csv"

# Initialize configuration variables from config xml file
$TenantID = $ConfigParams.SelectSingleNode("tenantid").InnerText
$APPObjectID = $ConfigParams.SelectSingleNode("appid").InnerText

Connect-MgGraph -ClientId $APPObjectID -TenantId $TenantID -CertificateName "CN=O365Monitor" -ErrorAction SilentlyContinue -Errorvariable ConnectionError | Out-Null

if($ConnectionError -ne $null)
    {
    Write-Host "$ConnectionError" -Foregroundcolor Red
    Exit
    }

if (Test-Path -Path $LogFile -PathType Leaf){
    rm $LogFile
}

$NotificationFlag = 0
$Headers = '"ID","Group Name","Group Description"'
$GroupList = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All
$PublicGroupList = $GroupList | Where-Object -Property Visibility -eq "Public"

Add-Content $LogFile $Headers

foreach ($Group in $PublicGroupList)
{
    Update-MgGroup -GroupId $Group.Id -Visibility "Private"
    Write-Host -NoNewLine "."
    $Entry = '"' + $Group.Id + '","' + $Group.DisplayName + '","' + $Group.Description + '"'
    Add-Content $LogFile $Entry
    $NotificationFlag++
}

if($NotificationFlag){
    Write-Output "Results can be found in $LogFile"
}
else{
    Write-Output "Congratulations, no public M365 groups found."
}

Disconnect-MgGraph | Out-Null
