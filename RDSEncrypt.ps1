<# 
.SYNOPSIS 
RDSEncrypt.ps1 - SSL/TLS certificate for Remote Desktop Services.
 
.DESCRIPTION  
This script will create, renew and install a SSL/TLS Let's Encrypt certificate on a Remote Desktop Services.
 
.OUTPUTS 
Outputs event data to the event viewer.
 
.NOTES 
Written by: Dylan McCrimmon

.LICENCE
This Script is licenced under the MIT License. View https://github.com/DylanNet/RDSEncrypt for more infomation.

Copyright 2019 Dylan McCrimmon
 
.CHANGE LOG
V1.00, 05/11/2019 - Creation
#> 



####### Custom Variables ########
       $emailAddress = "abc@dylanm.co.uk"              # Email Address for Alerts
         $domainName = "abc.dylanm.co.uk"              # Public Facing Domain Name
              $alias = "abc.dylanm.co.uk-01"           # Alias
        $iisSiteName = "Default Web Site"              # IIS Site Name
           $certPath = "C:\Certs\"                     # Location to save the certifcate
   $ConnectionBroker = "$env:COMPUTERNAME"             # Connection Broker
$CertificatePassword = "g1Bb3Rish"                     # Certifcate Password
#################################

#
##
#

##### Environment Variables #####
$certName = "$domainName-$alias-$(get-date -format HH--dd-MM-yyyy)"
$pfxfile =  "$certPath" + "$certname.pfx"
$CertificatePasswordSS = ConvertTo-SecureString "$CertificatePassword" -AsPlainText -Force

$IISACMEModuleName = "ACMESharp.Providers.IIS"
$IISACMEMinimumVersion = "0.9.1.326"
$ACMEModuleName = "ACMESharp"
$ACMEMinimumVersion = "0.9.1.326"
$eventLogSource = "Lets Encrypt for Remote Desktop Services"
#################################

function Write-Event {
    
    Param(
 
        [parameter(Mandatory=$true,position=1)]
        $Message,

        [parameter(Mandatory=$true,position=0)]
        [ValidateSet(
            'Infomation',
            'Warning',
            'Error'
        )]
        $EntryType,

        [parameter(Mandatory=$true,position=0)]
        $EventID
        )

    Write-EventLog –LogName Application –Source $eventLogSource –EntryType Information –EventID $EventID –Message $Message -Category ""

}

# Check if event log source exists
if (!([System.Diagnostics.EventLog]::SourceExists("$eventLogSource") -eq $true)) {
    
    # Create the event log source
    New-EventLog –LogName Application –Source $eventLogSource

}


Write-Event -EntryType Infomation -EventID "20021" -Message "Checking if ACMESharp Module is installed and imported."
# Check if the ACMESharp Module is installed and loaded
if (((Get-InstalledModule -Name "$ACMEModuleName").Name) -match ("$ACMEModuleName")) {
    
    # Module is installed

    Write-Event -EntryType Infomation -EventID "20021" -Message "ACMESharp Module is installed."
    Write-Event -EntryType Infomation -EventID "20021" -Message "Checking if ACMESharp Module is imported."

    # Check if the module is imported
    if (!(((Get-Module -Name "$ACMEModuleName").Name) -match ("$ACMEModuleName"))) {

        # Module is not Imported

        Write-Event -EntryType Infomation -EventID "20021" -Message "ACMESharp Module is not imported."

        # Import the module
        Import-Module $ACMEModuleName

        Write-Event -EntryType Infomation -EventID "20021" -Message "ACMESharp Module has been imported."
    }

} else {
    
    # Module is not installed

    Write-Event -EntryType Warning -EventID "20021" -Message "ACMESharp Module is not installed."

    # Install the required Module
    Write-Event -EntryType Infomation -EventID "20021" -Message "Installing ACMESharp Module."
    Install-Module -Name $ACMEModuleName -MinimumVersion "$ACMEMinimumVersion" -Force
    
    # Check if the module is imported
    if (!(((Get-Module -Name "$ACMEModuleName").Name) -match ("$ACMEModuleName"))) {

        # Module is not imported

        Write-Event -EntryType Infomation -EventID "20021" -Message "ACMESharp Module is not imported."

        # Import the module
        Import-Module $ACMEModuleName

        Write-Event -EntryType Infomation -EventID "20021" -Message "ACMESharp Module has been imported."
    }

}

# Check if the IIS ACMESharp Module is installed and loaded
if (((Get-InstalledModule -Name "$IISACMEModuleName").Name) -match ("$IISACMEModuleName")) {
    
    # Module is installed

    # Check if the module is imported
    if (!(((Get-Module -Name "$IISACMEModuleName").Name) -match ("$IISACMEModuleName"))) {

        # Module is not Imported

        # Import the module
        Import-Module $IISACMEModuleName
        Enable-ACMEExtensionModule $IISACMEModuleName

    }

} else {
    
    # Module is not installed

    # Install the required Module
    Install-Module -Name $IISACMEModuleName -MinimumVersion "$IISACMEMinimumVersion" -Force
    
    # Check if the module is imported
    if (!(((Get-Module -Name "$IISACMEModuleName").Name) -match ("$IISACMEModuleName"))) {

        # Module is not imported

        # Import the module
        Import-Module $IISACMEModuleName
        Enable-ACMEExtensionModule $IISACMEModuleName
    }

}

# Initialize the ACMEVault
Write-Event -EntryType Infomation -EventID "20021" -Message "Initializing the ACMEVault."

Initialize-ACMEVault -ErrorAction SilentlyContinue

Write-Event -EntryType Infomation -EventID "20021" -Message "ACMEVault has been initialized."

# Check if the server/person is registered
Try {
    
    Write-Event -EntryType Infomation -EventID "20021" -Message "Checking ACME Registration."

    Get-ACMERegistration

} Catch {

    # Server/person is not registered
    Write-Event -EntryType Warning -EventID "20021" -Message "No Active ACME Registration."

    # Create a registration
    Write-Event -EntryType Infomation -EventID "20021" -Message "Creating a ACME Registration."

    New-ACMERegistration -Contacts "mailto:$emailAddress" -AcceptTos

    # Send to LE
    Write-Event -EntryType Infomation -EventID "20021" -Message "Updating ACME Registration."

    Update-ACMERegistration

}



# Check if there is an Identifier
Try {

    Write-Event -EntryType Infomation -EventID "20021" -Message "Checking if there is an ACME identifier."

    Get-ACMEIdentifier $alias -ErrorAction Stop

} Catch {
    
    Write-Event -EntryType Warning -EventID "20021" -Message "No Active ACME Identifier."
    Write-Event -EntryType Infomation -EventID "20021" -Message "Creating an ACME Identifier."

    New-ACMEIdentifier -Dns $domainName -Alias $alias

    # Completing a challenge via http
    Write-Event -EntryType Infomation -EventID "20021" -Message "Completing the Let's Encrypt domain verification via http-01 challenge."

    Complete-ACMEChallenge $alias -ChallengeType http-01 -Handler iis -HandlerParameters @{ WebSiteRef = "$iisSiteName" } -Force

    # Submitting the challenge
    Write-Event -EntryType Infomation -EventID "20021" -Message "Submitting the domain verification to Let's Encrypt."

    Submit-ACMEChallenge $alias -ChallengeType http-01 -Force

    # Update the Identifier
    Write-Event -EntryType Infomation -EventID "20021" -Message "Updating the ACME Identifier."

    Update-ACMEIdentifier $alias -ChallengeType http-01

}

Write-Event -EntryType Infomation -EventID "20021" -Message "Checking the status of the domain with Let’s Encrypt."
# Check the status of the domain with Let’s Encrypt
if ((Update-ACMEIdentifier $alias).Status -eq "pending") {

    Write-Event -EntryType Warning -EventID "20021" -Message "Domain verification is pending."

    # Completing a challenge via http
    Write-Event -EntryType Infomation -EventID "20021" -Message "Completing the Let's Encrypt domain verification via http-01 challenge."

    Complete-ACMEChallenge $alias -ChallengeType http-01 -Handler iis -HandlerParameters @{ WebSiteRef = "$iisSiteName" } -Force

    # Submitting the challenge
    Write-Event -EntryType Infomation -EventID "20021" -Message "Submitting the domain verification to Let's Encrypt."

    Submit-ACMEChallenge $alias -ChallengeType http-01 -Force

    # Update the Identifier
    Write-Event -EntryType Infomation -EventID "20021" -Message "Updating the ACME Identifier."

    Update-ACMEIdentifier $alias -ChallengeType http-01

} elseif ((Update-ACMEIdentifier $alias).Status -eq 'valid') {
    
    # Creating a request to Let’s Encrypt
    Write-Event -EntryType Infomation -EventID "20021" -Message "Creating a request for a SSL/TLS certifcate."

    New-ACMECertificate $alias -Generate -Alias $certName

    # Submit the certifcate
    Write-Event -EntryType Infomation -EventID "20021" -Message "Sending the request for a SSL/TLS certifcate to Let’s Encrypt."

    Submit-ACMECertificate $certName -Force

    # Request for the certifcate to be updated
    Write-Event -EntryType Infomation -EventID "20021" -Message "Updating the certifcate."

    Update-ACMECertificate $certName

    # Get the SSL certifcate
    Write-Event -EntryType Infomation -EventID "20021" -Message "Downloading the SSL/TLS certifcate from Let’s Encrypt."

    Get-ACMECertificate $certName -ExportPkcs12 "$pfxfile" -CertificatePassword "$CertificatePassword"

    # Install the SSL certifcate to the RDPublishing
    Write-Event -EntryType Infomation -EventID "20022" -Message "Install the SSL/TLS certifcate to the RDPublishing."

    Set-RDCertificate -Role RDPublishing -ImportPath $pfxfile -Password $CertificatePasswordSS -ConnectionBroker $ConnectionBroker -Force

    # Install the SSL certifcate to the RDWebAccess
    Write-Event -EntryType Infomation -EventID "20022" -Message "Install the SSL/TLS certifcate to the RDWebAccess."

    Set-RDCertificate -Role RDWebAccess -ImportPath $pfxfile -Password $CertificatePasswordSS -ConnectionBroker $ConnectionBroker -Force

    # Install the SSL certifcate to the RDRedirector
    Write-Event -EntryType Infomation -EventID "20022" -Message "Install the SSL/TLS certifcate to the RDRedirector."

    Set-RDCertificate -Role RDRedirector -ImportPath $pfxfile -Password $CertificatePasswordSS -ConnectionBroker $ConnectionBroker -Force

    # Install the SSL certifcate to the RDGateway
    Write-Event -EntryType Infomation -EventID "20022" -Message "Install the SSL/TLS certifcate to the RDGateway."

    Set-RDCertificate -Role RDGateway -ImportPath $pfxfile -Password $CertificatePasswordSS -ConnectionBroker $ConnectionBroker -Force

}