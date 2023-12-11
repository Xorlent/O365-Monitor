#Requires -RunAsAdministrator
<#
.SYNOPSIS
Renews the self-signed certificate on a system already configured with the O365 Monitor scripts.  This process needs to be done yearly.
.DESCRIPTION
.EXAMPLE
.\Renew-O365MonitorCert.ps1
This will:
1. Create a new self signed certificate with the common name "CN=O365Monitor"
2. Uploads the generated public certificate for authentication to the O365 Monitor Enterprise App
#>

$CommonName = 'O365Monitor'
$StartDate = (Get-Date).ToString("yyyy-MM-dd")
$EndDate = (Get-Date).AddDays(385).ToString("yyyy-MM-dd")

function CreateSelfSignedCertificate()
{
    $certs = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object{$_.Subject -eq "CN=$CommonName"}
    if($certs -ne $null -and $certs.Length -gt 0)
    {
        foreach($c in $certs)
        {
            remove-item $c.PSPath
        }
    }

    $name = new-object -com "X509Enrollment.CX500DistinguishedName.1"
    $name.Encode("CN=$CommonName", 0)

    $key = new-object -com "X509Enrollment.CX509PrivateKey.1"
    $key.ProviderName = "Microsoft RSA SChannel Cryptographic Provider"
    $key.KeySpec = 1
    $key.Length = 2048
    $key.SecurityDescriptor = "D:PAI(A;;0xd01f01ff;;;SY)(A;;0xd01f01ff;;;BA)(A;;0x80120089;;;NS)"
    $key.MachineContext = 1
    $key.ExportPolicy = 1 # This is required to allow the private key to be exported
    $key.Create()

    $serverauthoid = new-object -com "X509Enrollment.CObjectId.1"
    $serverauthoid.InitializeFromValue("1.3.6.1.5.5.7.3.1") # Server Authentication
    $ekuoids = new-object -com "X509Enrollment.CObjectIds.1"
    $ekuoids.add($serverauthoid)
    $ekuext = new-object -com "X509Enrollment.CX509ExtensionEnhancedKeyUsage.1"
    $ekuext.InitializeEncode($ekuoids)

    $cert = new-object -com "X509Enrollment.CX509CertificateRequestCertificate.1"
    $cert.InitializeFromPrivateKey(2, $key, "")
    $cert.Subject = $name
    $cert.Issuer = $cert.Subject
    $cert.NotBefore = $StartDate
    $cert.NotAfter = $EndDate
    $cert.X509Extensions.Add($ekuext)
    $cert.Encode()

    $enrollment = new-object -com "X509Enrollment.CX509Enrollment.1"
    $enrollment.InitializeFromRequest($cert)
    $certdata = $enrollment.CreateRequest(0)
    $enrollment.InstallResponse(2, $certdata, 0, "")
    return $true
}

Function ConnectMgGraphModule
{
    $MsGraphModule =  Get-Module Microsoft.Graph -ListAvailable
    if($MsGraphModule -eq $null)
    { 
        Write-host "Important: Microsoft graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
        $confirm = Read-Host Are you sure you want to install Microsoft graph module? [Y] Yes [N] No  
        if($confirm -match "[yY]") 
        { 
            Write-host "Installing Microsoft graph module..."
            Install-Module Microsoft.Graph -Scope CurrentUser
            Write-host "Microsoft graph module installed successfully." -ForegroundColor Magenta 
        } 
        else
        { 
            Write-host "Exiting. `nNote: Microsoft graph module unavailable.  Exiting." -ForegroundColor Red
            Exit 
        } 
    }
    Connect-MgGraph -Scopes "Application.ReadWrite.All,Directory.ReadWrite.All"  -ErrorAction SilentlyContinue -Errorvariable ConnectionError |Out-Null
    if($ConnectionError -ne $null)
    {
        Write-Host "$ConnectionError" -Foregroundcolor Red
        Exit
    }
    Write-Host "Connected to O365 via Microsoft Graph..." -ForegroundColor Green
    $Script:TenantID = (Get-MgOrganization).Id
}

function UploadCertificate 
{
    Write-Progress -Activity "Uploading certificate..."
    $cert = Get-ChildItem -Path Cert:\LocalMachine\My | where-object{$_.Subject -eq "CN=$CommonName"}
    $KeyCredential = @{
        Type  = "AsymmetricX509Cert";
        Usage = "Verify";
        key   = $cert.RawData
    }
    Update-MgApplication -ApplicationId $APPObjectID  -KeyCredentials $KeyCredential -ErrorAction SilentlyContinue -ErrorVariable ApplicationError
    if($ApplicationError -ne $null)
    {
        Write-Host "$ApplicationError" -ForegroundColor Red
        CloseConnection 
    }
    Write-Host "`nCertificate uploaded successfully." -ForegroundColor Green
    $Script:Thumbprint = $Certificate.Thumbprint
}

CreateSelfSignedCertificate
ConnectMgGraphModule
UploadCertificate

Write-Output "Certificate Renewal Complete.  Next step, configure certificate private key permissions for your O365 Monitor service account."
Write-Output "Details here: https://github.com/Xorlent/O365-Monitor"