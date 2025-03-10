$ConfigFile = "$PSScriptRoot\O365Monitor-Config.xml"
$xml = New-Object System.Xml.XmlDocument
$xml.Load($ConfigFile)
$ConfigParams = $xml.SelectSingleNode("//o365app")

$LogFile = '$PSScriptRoot\O365Montior-DormantAccounts.csv'

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
$Headers = '"Last Sign-In","Display Name","User Principal Name"'
$UserList = Get-MgUser -Filter 'accountEnabled eq true' -All

Add-Content $LogFile $Headers

foreach ($User in $UserList)
{
    $UserID = $User.Id
    $UserName = $User.DisplayName
    $UserEmail = $User.UserPrincipalName
    $LastLogin = Get-MgUser -UserId $UserID -Property 'SignInActivity'
    Write-Host -NoNewLine "."
    if($LastLogin.SignInActivity.LastSignInDateTime -lt (Get-Date).AddDays(-45) -and $User.CreatedDateTime -lt (Get-Date).AddDays(-60)){
        $Entry = '"' + $LastLogin.SignInActivity.LastSignInDateTime + '","' + $UserName + '","' + $UserEmail + '"'
        Add-Content $LogFile $Entry
        Write-Host -NoNewLine "!"
        $NotificationFlag++
    }
    Start-Sleep -Seconds 1.5 # Throttle sign-in activity API requests
}

if($NotificationFlag){
    Write-Output "Results can be found in $LogFile"
}
else{
    Write-Output "No dormant accounts found."
}

Disconnect-MgGraph | Out-Null
