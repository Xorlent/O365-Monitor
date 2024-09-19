#Requires -RunAsAdministrator
<#
.SYNOPSIS
Prepares a host to run the O365 Monitor scripts, configures the O365 Monitor Enterprise App in Office 365
.DESCRIPTION
.EXAMPLE
.\Install-O365Monitor.ps1
This will:
1. Create a new self-signed certificate with the common name "CN=O365Monitor"
2. Create an Enterprise App, "O365 Monitor" with only the permissions necessary to conduct its scans:
    Application.Read.All
    User.Read.All
    AuditLog.Read.All
3. Uploads the generated public certificate for authentication to the newly created "O365 Monitor" Enterprise App
4. Prompts for a tenant admin to grant the configured Enterprise App permissions
.EXAMPLE
.\Install-O365Monitor.ps1 -Force
This will first delete any preexisting O365 Monitor certificate and creates a new self-signed certificate with the common name "CN=O365Monitor"
#>
Param(
   [Parameter(Mandatory=$false, HelpMessage="Will overwrite existing certificates")]
   [Switch]$Force
)

$ConfigFile = $PWD.Path + '\O365Monitor-Config.xml'
$CommonName = 'O365Monitor'
$StartDate = (Get-Date).ToString("yyyy-MM-dd")
$EndDate = (Get-Date).AddDays(385).ToString("yyyy-MM-dd")

function CreateSelfSignedCertificate()
{
    $certs = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object{$_.Subject -eq "CN=$CommonName"}
    if($certs -ne $null -and $certs.Length -gt 0)
    {
        if($Force)
        {

            foreach($c in $certs)
            {
                remove-item $c.PSPath
            }
        } else {
            Write-Host -ForegroundColor Red "One or more certificates with the same common name (CN=$CommonName) are already located in the local certificate store. Use -Force to remove them";
            return $false
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
        Write-host 'If completion hangs or fails, please manually run "Update-Module" in a PowerShell window to first update prerequisite components.' -ForegroundColor Yellow
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

function RegisterApplication
{
    Write-Progress -Activity "Registering O365 Monitor Enterprise App..."

    $Script:RedirectURI = "https://login.microsoftonline.com/common/oauth2/nativeclient"
    $params = @{
        DisplayName = "O365 Monitor"
        SignInAudience="AzureADMyOrg"
        Notes="Auto-created by O365 Monitor PowerShell App"
        PublicClient=@{
                RedirectUris = "$RedirectURI"
        }
        RequiredResourceAccess = @(
            @{
                ResourceAppId = "00000003-0000-0000-c000-000000000000" #  Microsoft Graph Resource ID
                ResourceAccess = @(
                    @{
                        Id = "9a5d68dd-52b0-4cc2-bd40-abcf44ac3a30" #Application.Read.All
                        Type = "Role"                               #Role -> Application permission
                     }
                    @{
                        Id = "a154be20-db9c-4678-8ab7-66f6cc099a59" #User.Read.All
                        Type = "Role"                               #Role -> Application permission
                     }
                    @{
                        Id = "b0afded3-3588-46d8-8b3d-9842eff778da" #AuditLog.Read.All
                        Type = "Role"                               #Role -> Application permission
                     }
                    @{
                        Id = "62a82d76-70ea-41e2-9197-370581804d09" #Group.ReadWrite.All
                        Type = "Role"                               #Role -> Application permission
                     }
                    )
                }
            )
    }
    try{
        $Script:App = New-MgApplication -BodyParameter $params
    }
    catch
    {
        Write-Host $_.Exception.Message -ForegroundColor Red
        CloseConnection
    }
    
    $Script:APPObjectID = $App.Id
    $Script:APPID = $App.AppId
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

function GrantPermission
{
    Write-Progress -Activity "Granting admin consent..."
    Start-Sleep -Seconds 20
    $Script:ClientID = $App.AppId
    $URL = "https://login.microsoftonline.com/$TenantID/adminconsent"
    $Url="$URL`?client_id=$ClientID"
    Write-Host "`nMS Graph requires admin consent to access data. Please grant access to the application." -ForegroundColor Cyan
    While(1)
    {
        Add-Type -AssemblyName System.Windows.Forms
        $script:mainForm = New-Object System.Windows.Forms.Form -Property @{
            Width  = 680
            Height = 640
        }
        $script:webBrowser = New-Object System.Windows.Forms.WebBrowser -Property @{
            Width  = 680
            Height = 640
            URL    = $URL
        }
        $document={
            if($webBrowser.Url -eq "$RedirectURI`?admin_consent=True&tenant=$TenantID" -or $webBrowser.Url -match "error")
            {
                $mainForm.Close()
            }
            if($webBrowser.DocumentText.Contains("We received a bad request"))
            {
                $mainForm.Close()
            }
        }
        $webBrowser.ScriptErrorsSuppressed = $true
        $webBrowser.Add_DocumentCompleted($document)
        $mainForm.Controls.Add($webBrowser)
        $mainForm.Add_Shown({ $mainForm.Activate() ;$mainForm.Refresh()})
        [void] $mainForm.ShowDialog()
        if($webBrowser.Url.AbsoluteUri -eq "$RedirectURI`?admin_consent=True&tenant=$TenantID")
        {
            Write-Host "`nAdmin consent granted successfully." -ForegroundColor Green
            break
        }
        else
        {
            Write-Host "`nAdmin consent failed." -ForegroundColor Red
            $Confirm = Read-Host "Do you want to retry admin consent? [Y] Yes [N] No"
            if($Confirm -match "[yY]")
            {
                Continue
            } 
            else
            {
                Write-Host "You can manually grant admin consent in the Azure AD/Entra portal." -ForegroundColor Yellow
                break
            }
        }
    }
}

CreateSelfSignedCertificate
ConnectMgGraphModule
RegisterApplication
UploadCertificate
GrantPermission
Disconnect-MgGraph | Out-Null

# If the config file exists, rename it first.
if (Test-Path -Path $ConfigFile -PathType Leaf){
    Remove-Item -Path ".\$ConfigFile.bak" -Force | Out-Null
    Rename-Item -Path $ConfigFile -NewName "$ConfigFile.bak"
    }

# Save the the TenantID and ApplicationID configuration values to XML file
$xml = New-Object System.Xml.XmlDocument
$xml.AppendChild($xml.CreateXmlDeclaration("1.0", "UTF-8", $null))
$root = $xml.AppendChild($xml.CreateElement("o365app"))
$root.AppendChild($xml.CreateElement("tenantid")).InnerText = $TenantID
$root.AppendChild($xml.CreateElement("appid")).InnerText = $APPObjectID
$xml.Save($ConfigFile)

Write-Output "Setup Complete.  Next step, configure certificate private key permissions for your O365 Monitor service account."
Write-Output "Details here: https://github.com/Xorlent/O365-Monitor"
