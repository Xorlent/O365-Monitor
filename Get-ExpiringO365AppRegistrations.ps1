$ConfigFile = '.\O365Monitor-Config.xml'
$ConfigParams = [xml](get-content $ConfigFile)

$LogFile = '.\O365Montior-ExpiringCerts.csv'

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

$Now = Get-Date
$Applications = Get-MgApplication -all
$Logs = @()
$NotificationFlag = 0

foreach ($App in $Applications) {
    $AppName = $App.DisplayName
    $AppID   = $App.Id
    $ApplID  = $App.AppId

    $AppCreds = Get-MgApplication -ApplicationId $AppID | Select-Object PasswordCredentials, KeyCredentials

    $Secrets = $AppCreds.PasswordCredentials
    $Certs   = $AppCreds.KeyCredentials

    foreach ($Secret in $Secrets) {
        $StartDate  = $Secret.StartDateTime
        $EndDate    = $Secret.EndDateTime
        $SecretName = $Secret.DisplayName

        $RemainingDaysCount = ($EndDate - $Now).Days

        if ($RemainingDaysCount -le 45) {
            $Logs += [PSCustomObject]@{
                'App Name'        = $AppName
                'App ID'          = $ApplID
                'Secret Name'            = $SecretName
                'Secret Expiration'      = $EndDate
                'Certificate Name'       = $Null
                'Certificate Expiration' = $Null
            }
            $NotificationFlag++
        }
    }

    foreach ($Cert in $Certs) {
        $StartDate = $Cert.StartDateTime
        $EndDate   = $Cert.EndDateTime
        $CertName  = $Cert.DisplayName

        $RemainingDaysCount = ($EndDate - $Now).Days

        if ($RemainingDaysCount -le 45) {
            $Logs += [PSCustomObject]@{
                'App Name'        = $AppName
                'App ID'          = $ApplID
                'Secret Name'            = $Null
                'Certificate Name'       = $CertName
                'Certificate Expiration' = $EndDate
            }
            $NotificationFlag++
        }
    }
}

if($NotificationFlag){
    $Logs | Out-File -FilePath $LogFile
    }

Disconnect-MgGraph
