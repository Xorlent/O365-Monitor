$ConfigFile = '.\O365Monitor-Config.xml'
$ConfigParams = [xml](get-content $ConfigFile)

$LogFile = '.\O365Montior-DormantUsers.csv'

# Initialize configuration variables from config xml file
$TenantID = $ConfigParams.o365app.tenantid.value
$APPObjectID = $ConfigParams.o365app.appid.value

Connect-MgGraph -ClientID $APPObjectID -TenantId $TenantID -CertificateName "CN=O365Monitor" -ErrorAction SilentlyContinue -Errorvariable ConnectionError | Out-Null

if($ConnectionError -ne $null)
    {
    Write-Host "$ConnectionError" -Foregroundcolor Red
    Exit
    }

if (Test-Path -Path $LogFile -PathType Leaf){
    rm $LogFile
}

$Headers = '"Last Sign-In","Display Name","User Principal Name"'
$UserList = Get-MgUser -Filter 'accountEnabled eq true' -All

Add-Content $LogFile $Headers

foreach ($User in $UserList)
{
    $UserID = $User.Id
    $UserName = $User.DisplayName
    $UserEmail = $User.UserPrincipalName
    $LastLogin = Get-MgUser -UserId $UserID -Property 'SignInActivity'

    if($LastLogin.SignInActivity.LastSignInDateTime -lt (Get-Date).AddDays(-45)){
        $Entry = '"' + $LastLogin.SignInActivity.LastSignInDateTime + '","' + $UserName + '","' + $UserEmail + '"'
        Add-Content $LogFile $Entry
    }
}

Disconnect-MgGraph
