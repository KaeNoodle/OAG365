<#
.SYNOPSIS
OAG Defender for Office 365 Policy Report

.DESCRIPTION
Uses the Microsoft.Graph PowerShell modules to export Defender for Office 365 policy settings

   Export: Exchange Online Protection and Defender for Office threat protection policies
    Roles: The user must authenticate interactively and have one of the following roles.
             Entra: Global Admin, Global Reader*, Security Administrator, Security Reader*; OR
             Exchange Online: Security Administrator, Organization Management, View-Only Organization Management*, Global Reader*, Security Reader*

    * indicates ready-only, least privilege rights

.COMPONENT 
Requires PowerShell module ExchangeOnline
Specific requirements documented in $script:psExoModulesRequired

.INPUTS
Prompted for EntraID credentials with the roles itemised above

.OUTPUTS
CSV exports containing
  - Exchange Online Protection and Defender Threat Protection policies

.PARAMETER Output
The name of the folder to output exported information to. Defaults to the running scripts folder ($PSScriptRoot).

.EXAMPLE
Export-M365-DfoReport.ps1

By default the script will prompt for interactive login as a user with the required roles. Results will be saved in a subfolder under the scripts current location. 

.EXAMPLE
Export-M365-DfoReport.ps1 -Output C:\temp

Optionally you can change the location the exported files are saved to. E.g. C:\temp\M365\YYYYMMDD_HHIISS\*


.NOTES
VERSION: 3.0

RELEASE APPROVAL: <UNAPPROVED>

2026-05-12
  - Code clean-up. Updated to use approved verbs and clearer function names

2026-05-01: Created as standalone script

Uses ExchangeOnlineManagement PowerShell module:
    # Global
    Connect-ExchangeOnline
    Disconnect-ExchangeOnline

    # Exchange Online Protection and Defender Threat Policies
    Get-MalwareFilterPolicy
    Get-MalwareFilterRule
    Get-AntiPhishingPolicy
    Get-AntiPhishingRule
    Get-HostedConnectionFilterPolicy
    Get-SafeAttachmentsPolicy
    Get-SafeAttachmentsRule
    Get-SafeLinksPolicy
    Get-SafeLinksRule
    Get-QuarantinePolicy

.LINK

#>
Param (
	[String]$output = $PSScriptRoot
)

begin {

function Format-Oag365-Exception {
    <#
    .SYNOPSIS
    Exception message handling. 

    .DESCRIPTION
    Create and return a formatted string containing and error message generated from the provided exceptions object.  

    .PARAMETER message
    Custom error message provided by the application to provide meaningful context 

    .PARAMETER exception
    Exception object as caught by try / catch

    .EXAMPLE
    Format-Oag365-Exception -message "Failed connecting Graph API to M365 tenant" -exception $e

    .NOTES
        NAME: Format-Oag365-Exception
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:

        CHANGELOG:
    #>

    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)][string]$message,
        [Parameter(Mandatory = $true)][AllowEmptyString()][object]$exception
    )

    $fmtString = "{0}`n`n{1} : {2}`n{3}`n" +
                 "      + CategoryInfo     : {4}`n" +
                 "      + FullyQualifiedId : {5}`n"
    
    $exceptionFields = $message, $exception.InvocationInfo.InvocationName, $exception.Exception.Message, $exception.InvocationInfo.PositionMessage, $exception.CategoryInfo.ToString(), $exception.FullyQualifiedErrorId

    return $fmtString -f $exceptionFields
}

function Connect-Oag365Dfo-ToExchangeOnline {
    <#
    .SYNOPSIS
    Connects PowerShell to M365 Exchange Online Management  

    .DESCRIPTION
    Imports required PowerShell modules, if required, and then connects the Exchange Online Management PowerShell module.  

    .EXAMPLE
    Connect-Oag365Dfo-ToExchangeOnline

    .NOTES
        NAME: Connect-Oag365Dfo-ToExchangeOnline
        VERSION: 1.1

        FUNCTIONS & PERMISSIONS:
           Connect-ExchangeOnline

        CHANGELOG:    
          2025-11-18: Updated in include modules only when required. Part of allowing stages to be run independently.
    #>

    try {
        Write-Host "`nMicrosoft Exchange Online Management PowerShell required. Please login if prompted" -ForegroundColor Cyan
        Start-Sleep 2

        if (-not (Get-Module -Name $script:psExoModulesRequired)) {
            Write-Host " - Importing PowerShell Exchange Online modules..." -ForegroundColor Cyan
            Start-Sleep 2
            $script:psExoModulesRequired | Import-Module -Verbose
        }

        Write-Host " - Connecting to Microsoft Exchange Online PowerShell..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowProgress $true -ShowBanner:$false

        $connection = Get-ConnectionInformation
        if ($connection.state -eq "Connected") {
            Write-Host "     - State: $($connection.state)" -ForegroundColor Green
        } else {
            Write-Host "     - State: $($connection.state)" -ForegroundColor Red
        }
        Write-Host "     - User: $($connection.UserPrincipalName)" -ForegroundColor Cyan
        Write-Host "`n"
        Pause

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed connecting to Microsoft Exchange Online PowerShell" -exception $_ 
        Write-Host $exception -ForegroundColor Red
        pause
        return $false
    }
}

function Disconnect-Oag365Dfo-FromExchangeOnline {
    <#
    .SYNOPSIS
    Disconnects from M365 services

    .DESCRIPTION
    Disconnects from Connect-MgGraph and Connect-ExcangeOnline.

    .EXAMPLE
    Disconnect-Oag365Dfo-FromExchangeOnline

    .NOTES
        NAME: Disconnect-Oag365Dfo-FromExchangeOnline
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Disconnect-ExchangeOnline

        CHANGELOG:    
    #>

    try {
        Disconnect-ExchangeOnline -ErrorAction SilentlyContinue -Confirm:$false  
    } catch {
        Write-Host "ERROR: Failed disconnecting $($_.Exception.message)" -ForegroundColor Red
    }
}




function Export-Oag365Dfo-AntiMalwareConfig {
    <#
    .SYNOPSIS
    Export anti-malware policy and rule threat protection policies

    .DESCRIPTION
    Obtain anti-malware policy and rule settings from Exchange Online / Defender threat protection policies (Defender for Office 365)

    .EXAMPLE
    Export-Oag365Dfo-AntiMalwareConfig

    .NOTES
        NAME: Export-Oag365Dfo-AntiMalwareConfig
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-MalwareFilterPolicy
           Get-MalwareFilterRule

        CHANGELOG:
    #>

    try {
        Write-Host "`nExchange Online Protection anti-malware protection settings"

        # Get EOP malware policy settings
        $getMalwareFilterPolicy = Get-MalwareFilterPolicy 
        $getMalwareFilterPolicy | Export-CSV $script:exportTarget.eopMalwarePolicy -NoTypeInformation
        Write-Host " - Exported EOP anti-malware protection policy settings to $($script:exportTarget.eopMalwarePolicy)" -ForegroundColor Green

        $getMalwareFilterRule = Get-MalwareFilterRule 
        $getMalwareFilterRule | Export-CSV $script:exportTarget.eopMalwareRule -NoTypeInformation
        Write-Host " - Exported EOP anti-malware protection rules to $($script:exportTarget.eopMalwareRule)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving EOP anti-malware policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function Export-Oag365Dfo-AntiPhishingConfig {
    <#
    .SYNOPSIS
    Export anti-malware policy and rule threat protection policies

    .DESCRIPTION
    Export anti-malware policy and rule settings from Exchange Online / Defender threat protection policies (Defender for Office 365)

    .EXAMPLE
    Export-Oag365Dfo-AntiPhishingConfig

    .NOTES
        NAME: Export-Oag365Dfo-AntiPhishingConfig
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-AntiPhishPolicy
           Get-AntiPhishRule

        CHANGELOG:
    #>

    try {
        Write-Host "`nExchange Online Protection anti-phishing protection settings"

        # Get EOP phishing policy settings
        $getAntiPhishPolicy = Get-AntiPhishPolicy 
        $getAntiPhishPolicy | Export-CSV $script:exportTarget.eopPhishPolicy -NoTypeInformation
        Write-Host " - Exported EOP anti-phishing protection policy settings to $($script:exportTarget.eopPhishPolicy)" -ForegroundColor Green

        # Export to CSV
        $getAntiPhishRule = Get-AntiPhishRule 
        $getAntiPhishRule | Export-CSV $script:exportTarget.eopPhishRule -NoTypeInformation
        Write-Host " - Exported EOP anti-phishing protection rule to $($script:exportTarget.eopPhishRule)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving EOP anti-phishing policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function Export-Oag365Dfo-AntiSpamConfig {
    <#
    .SYNOPSIS
    Export anti-spam policy threat protection policies

    .DESCRIPTION
    Export anti-spam policy settings from Exchange Online / Defender Threat Protection (Defender for Office 365)

    .EXAMPLE
    Export-Oag365Dfo-AntiSpamConfig

    .NOTES
        NAME: Export-Oag365Dfo-AntiSpamConfig
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-HostedConnectionFilterPolicy

        CHANGELOG:
    #>

    try {
        Write-Host "`nExchange Online Protection anti-spam protection settings"

        # Get EOP spam policy settings
        $getHostedConnectionFilterPolicy = Get-HostedConnectionFilterPolicy 
        $getHostedConnectionFilterPolicy | Export-CSV $script:exportTarget.eopSpamPolicy -NoTypeInformation
        Write-Host " - Exported EOP anti-spam protection policy settings to $($script:exportTarget.eopSpamPolicy)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving EOP anti-spam policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function Export-Oag365Dfo-SafeLinksConfig {
    <#
    .SYNOPSIS
    Export safe links policy threat protection policies

    .DESCRIPTION
    Export safe links policy settings from Defender Threat Protection (Defender for Office 365)

    .EXAMPLE
    Export-Oag365Dfo-SafeLinksConfig

    .NOTES
        NAME: Export-Oag365Dfo-SafeLinksConfig
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-SafeLinksPolicy
           Get-SafeLinksRule

        CHANGELOG:
    #>

    try {
        Write-Host "`nDefender Threat Protection safe links policy settings"

        # Get DTP safe links settings
        $getSafeLinksPolicy = Get-SafeLinksPolicy 
        $getSafeLinksPolicy | Export-CSV $script:exportTarget.dtppSafeLinksPolicy -NoTypeInformation
        Write-Host " - Exported DTP safe links policy settings to $($script:exportTarget.dtppSafeLinksPolicy)" -ForegroundColor Green

        $getSafeLinksRule = Get-SafeLinksRule 
        $getSafeLinksRule | Export-CSV $script:exportTarget.dtppSafeLinksRule -NoTypeInformation
        Write-Host " - Exported DTP safe links rule settings to $($script:exportTarget.dtppSafeLinksRule)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving DTP safe links policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function Export-Oag365Dfo-SafeAttachmentsConfig {
    <#
    .SYNOPSIS
    Export safe attachment threat protection policies

    .DESCRIPTION
    Export safe attachments policy settings from Defender threat protection policies (Defender for Office 365)

    .EXAMPLE
    Export-Oag365Dfo-SafeAttachmentsConfig

    .NOTES
        NAME: Export-Oag365Dfo-SafeAttachmentsConfig
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-SafeAttachmentPolicy
           Get-SafeAttachmentRule

        CHANGELOG:
    #>

    try {
        Write-Host "`nDefender Threat Protection safe attachments policy settings"

        # Get DTP safe attachments settings
        $getSafeAttachmentPolicy = Get-SafeAttachmentPolicy
        $getSafeAttachmentPolicy | Export-CSV $script:exportTarget.dtppSafeAttachmentPolicy -NoTypeInformation
        Write-Host " - Exported DTP safe attachments policy settings to $($script:exportTarget.dtppSafeAttachmentPolicy)" -ForegroundColor Green

        # Export as CSV
        $getSafeAttachmentRule = Get-SafeAttachmentRule
        $getSafeAttachmentRule | Export-CSV $script:exportTarget.dtppSafeAttachmentRule -NoTypeInformation
        Write-Host " - Exported DTP safe attachments rule to $($script:exportTarget.dtppSafeAttachmentRule)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving DTP safe attachments policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}

function Export-Oag365Dfo-QuarantinePolicy {
    <#
    .SYNOPSIS
    Export threat protection quarantine policies

    .DESCRIPTION
    Export quarantine policy settings from Exchange Online / Defender threat protection policies (Defender for Office 365)

    .EXAMPLE
    Export-Oag365Dfo-QuarantinePolicy

    .NOTES
        NAME: Export-Oag365Dfo-QuarantinePolicy
        VERSION: 1.0

        FUNCTIONS & PERMISSIONS:
           Get-QuarantinePolicy

        CHANGELOG:
    #>

    try {
        Write-Host "`nExchange Online quarantine policy settings"

        $getQuarantinePolicy = Get-QuarantinePolicy | Select-Object -Property Id, IsValid, Name, EndUserQuarantinePermissions, ESNEnabled, QuarantinePolicyType, QuarantineRetentionDays, IncludeMessagesFromBlockedSenderAddress, DistinguishedName, ObjectCategory, ObjectClass, WhenChanged, WhenCreated, WhenChangedUTC, WhenCreatedUTC, ExchangeObjectId, OrganizationalUnitRoot, OrganizationId, Guid, OriginatingServer, ObjectState, ExchangeVersion
        $getQuarantinePolicy | Export-CSV $script:exportTarget.dfoQuarantinePolicy -NoTypeInformation
        Write-Host " - Exported EOP quarantine policy settings to $($script:exportTarget.dfoQuarantinePolicy)" -ForegroundColor Green

        return $true
    } catch {
        $exception = Format-Oag365-Exception -message "Failed retrieving EOP quarantine policies" -exception $_
        Write-Host $exception -ForegroundColor Red
        return $false
    }
}




    # -----------------------------------
    # Define variables required by script
    # -----------------------------------

    Clear-Host

    $script:output = ($output | Resolve-Path).Path
    $script:execDateTime = (Get-Date).ToString("yyyyMMdd_HHmmss")

    # Create folders to store output
	$script:exportTargetFolder = [PSCustomObject]@{
		"threatProtection"  = New-Item -ItemType Directory -Force -Path "$($script:Output)\M365\$($script:execDateTime)\ThreatProtection"
	}

    # Define where exports will be saved
    $script:exportTarget = [pscustomobject]@{
        "dtppSafeLinksPolicy"       = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "dtppSafeLinksPolicy.csv"
        "dtppSafeLinksRule"         = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "dtppSafeLinksRule.csv"
        "dtppSafeAttachmentPolicy"  = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "dtppSafeAttachmentPolicy.csv"
        "dtppSafeAttachmentRule"    = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "dtppSafeAttachmentRule.csv"
        "eopMalwarePolicy"          = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "eopMalwarePolicy.csv"
        "eopMalwareRule"            = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "eopMalwareRule.csv"
        "eopPhishPolicy"            = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "eopPhishPolicy.csv"
        "eopPhishRule"              = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "eopPhishRule.csv"
        "eopSpamPolicy"             = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "eopSpamPolicy.csv"
        "dfoQuarantinePolicy"       = Join-Path -Path $script:exportTargetFolder.threatProtection -ChildPath "dfoQuarantinePolicy.csv"
    }

    $script:psExoModulesRequired = @(
        'ExchangeOnlineManagement'
    )



    # --------------------------------------------------------------------------------------------------------------
    # Initialise script. Start transcript, check for required modules and safe mode. Output parameters for audit log
    # --------------------------------------------------------------------------------------------------------------

    Start-Transcript -Path (Join-Path -Path $output -ChildPath "$($script:execDateTime)_ExportM365_transcript.txt")
    Write-Host "Starting $(Get-Date)"

    Write-Host "`n----------------------------------------------"
    Write-Host " This script will output audit evidence to a series of subfolders under the location '$output'"

    ($script:psExoModulesRequired) | ForEach-Object {
        $getModuleResult = Get-Module -ListAvailable -Name $_ -ErrorAction SilentlyContinue 
        if ($getModuleResult) {
            $moduleVersionNo = ($getModuleResult | Sort-Object Version -Descending).Version | Join-String -Separator ", "
            Write-Host "   - Found required PowerShell module '$_' [$moduleVersionNo]." -ForegroundColor Green
        } else {
            Write-Host "   - Missing required PowerShell module '$_'. Abort now." -ForegroundColor Red
            Write-Host "     Try: Install-Module $_ -Scope CurrentUser -Force -AllowClobber" -ForegroundColor Cyan
        }
    }

    Write-Host "   - If exceptions occur check that PowerShell modules are up-to-date." -ForegroundColor Cyan

    if ($ExecutionContext.SessionState.LanguageMode -eq "ConstrainedLanguage") {
	    Write-Host "   - ERROR. This script cannot run in Constrained Language Mode. Abort now." -ForegroundColor Red
    } elseif ($ExecutionContext.SessionState.LanguageMode -eq "FullLanguage") {
	    Write-Host "   - Running in Full Language Mode" -ForegroundColor Green
    } else {		
	    Write-Host "   - Caution'$($ExecutionContext.SessionState.LanguageMode)' Language Mode detected" -ForegroundColor DarkYellow
    }

	if ($PSVersionTable.PSVersion.Major -ge 7 -and $PSVersionTable.PSVersion.Minor -ge 1) {
		Write-Host "   - PowerShell version $($PSVersionTable.PSVersion) found, PowerShell 7.1+ is recommended." -ForegroundColor Green
	} else {
		Write-Host "   - Caution, PowerShell version '$($PSVersionTable.PSVersion)' detected. PowerShell 7.1+ is recommended. Some parts of the script may not operate correctly." -ForegroundColor DarkYellow
	}

    Write-Host "----------------------------------------------`n"
    Pause

    Write-Host "`nScript Parameters:" -ForegroundColor Cyan
    Write-Host "  Output Folder: $output"
    Write-Host "`n`n"
}

process {

    Write-Host "`nExport Exchange Online Protection and Defender Threat Protection Policies" -ForegroundColor Cyan
    
    if (Connect-Oag365Dfo-ToExchangeOnline -eq $true) {
        Export-Oag365Dfo-AntiMalwareConfig | Out-Null
        
        Export-Oag365Dfo-AntiPhishingConfig | Out-Null

        Export-Oag365Dfo-AntiSpamConfig | Out-Null

        Export-Oag365Dfo-SafeLinksConfig | Out-Null

        Export-Oag365Dfo-SafeAttachmentsConfig | Out-Null

        Export-Oag365Dfo-QuarantinePolicy | Out-Null
    }

}

end {
    Write-Host "`n`nCompleted $(Get-Date). Disconnecting from Microsoft Exchange Online..."
    Disconnect-Oag365Dfo-FromExchangeOnline

    Stop-Transcript
}


























#                                                                     
#                                                                     
#         OOOOOOOOO                 AAA                  GGGGGGGGGGGGG
#       OO:::::::::OO              A:::A              GGG::::::::::::G
#     OO:::::::::::::OO           A:::::A           GG:::::::::::::::G
#    O:::::::OOO:::::::O         A:::::::A         G:::::GGGGGGGG::::G
#    O::::::O   O::::::O        A:::::::::A       G:::::G       GGGGGG
#    O:::::O     O:::::O       A:::::A:::::A     G:::::G              
#    O:::::O     O:::::O      A:::::A A:::::A    G:::::G              
#    O:::::O     O:::::O     A:::::A   A:::::A   G:::::G    GGGGGGGGGG
#    O:::::O     O:::::O    A:::::A     A:::::A  G:::::G    G::::::::G
#    O:::::O     O:::::O   A:::::AAAAAAAAA:::::A G:::::G    GGGGG::::G
#    O:::::O     O:::::O  A:::::::::::::::::::::AG:::::G        G::::G
#    O::::::O   O::::::O A:::::AAAAAAAAAAAAA:::::AG:::::G       G::::G
#    O:::::::OOO:::::::OA:::::A             A:::::AG:::::GGGGGGGG::::G
#     OO:::::::::::::OOA:::::A               A:::::AGG:::::::::::::::G
#       OO:::::::::OO A:::::A                 A:::::A GGG::::::GGG:::G
#         OOOOOOOOO  AAAAAAA                   AAAAAAA   GGGGGG   GGBG
#                                                                     
#                                                                     

# SIG # Begin signature block
# MIIV5QYJKoZIhvcNAQcCoIIV1jCCFdICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDJcO+vZilr+DZ6
# P4fSGYETzBJ2JNdQN2h3aWzc7EHZmaCCEiAwggVvMIIEV6ADAgECAhBI/JO0YFWU
# jTanyYqJ1pQWMA0GCSqGSIb3DQEBDAUAMHsxCzAJBgNVBAYTAkdCMRswGQYDVQQI
# DBJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcMB1NhbGZvcmQxGjAYBgNVBAoM
# EUNvbW9kbyBDQSBMaW1pdGVkMSEwHwYDVQQDDBhBQUEgQ2VydGlmaWNhdGUgU2Vy
# dmljZXMwHhcNMjEwNTI1MDAwMDAwWhcNMjgxMjMxMjM1OTU5WjBWMQswCQYDVQQG
# EwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMS0wKwYDVQQDEyRTZWN0aWdv
# IFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQCN55QSIgQkdC7/FiMCkoq2rjaFrEfUI5ErPtx94jGgUW+s
# hJHjUoq14pbe0IdjJImK/+8Skzt9u7aKvb0Ffyeba2XTpQxpsbxJOZrxbW6q5KCD
# J9qaDStQ6Utbs7hkNqR+Sj2pcaths3OzPAsM79szV+W+NDfjlxtd/R8SPYIDdub7
# P2bSlDFp+m2zNKzBenjcklDyZMeqLQSrw2rq4C+np9xu1+j/2iGrQL+57g2extme
# me/G3h+pDHazJyCh1rr9gOcB0u/rgimVcI3/uxXP/tEPNqIuTzKQdEZrRzUTdwUz
# T2MuuC3hv2WnBGsY2HH6zAjybYmZELGt2z4s5KoYsMYHAXVn3m3pY2MeNn9pib6q
# RT5uWl+PoVvLnTCGMOgDs0DGDQ84zWeoU4j6uDBl+m/H5x2xg3RpPqzEaDux5mcz
# mrYI4IAFSEDu9oJkRqj1c7AGlfJsZZ+/VVscnFcax3hGfHCqlBuCF6yH6bbJDoEc
# QNYWFyn8XJwYK+pF9e+91WdPKF4F7pBMeufG9ND8+s0+MkYTIDaKBOq3qgdGnA2T
# OglmmVhcKaO5DKYwODzQRjY1fJy67sPV+Qp2+n4FG0DKkjXp1XrRtX8ArqmQqsV/
# AZwQsRb8zG4Y3G9i/qZQp7h7uJ0VP/4gDHXIIloTlRmQAOka1cKG8eOO7F/05QID
# AQABo4IBEjCCAQ4wHwYDVR0jBBgwFoAUoBEKIz6W8Qfs4q8p74Klf9AwpLQwHQYD
# VR0OBBYEFDLrkpr/NZZILyhAQnAgNpFcF4XmMA4GA1UdDwEB/wQEAwIBhjAPBgNV
# HRMBAf8EBTADAQH/MBMGA1UdJQQMMAoGCCsGAQUFBwMDMBsGA1UdIAQUMBIwBgYE
# VR0gADAIBgZngQwBBAEwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybC5jb21v
# ZG9jYS5jb20vQUFBQ2VydGlmaWNhdGVTZXJ2aWNlcy5jcmwwNAYIKwYBBQUHAQEE
# KDAmMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5jb21vZG9jYS5jb20wDQYJKoZI
# hvcNAQEMBQADggEBABK/oe+LdJqYRLhpRrWrJAoMpIpnuDqBv0WKfVIHqI0fTiGF
# OaNrXi0ghr8QuK55O1PNtPvYRL4G2VxjZ9RAFodEhnIq1jIV9RKDwvnhXRFAZ/ZC
# J3LFI+ICOBpMIOLbAffNRk8monxmwFE2tokCVMf8WPtsAO7+mKYulaEMUykfb9gZ
# pk+e96wJ6l2CxouvgKe9gUhShDHaMuwV5KZMPWw5c9QLhTkg4IUaaOGnSDip0TYl
# d8GNGRbFiExmfS9jzpjoad+sPKhdnckcW67Y8y90z7h+9teDnRGWYpquRRPaf9xH
# +9/DUp/mBlXpnYzyOmJRvOwkDynUWICE5EV7WtgwggYaMIIEAqADAgECAhBiHW0M
# UgGeO5B5FSCJIRwKMA0GCSqGSIb3DQEBDAUAMFYxCzAJBgNVBAYTAkdCMRgwFgYD
# VQQKEw9TZWN0aWdvIExpbWl0ZWQxLTArBgNVBAMTJFNlY3RpZ28gUHVibGljIENv
# ZGUgU2lnbmluZyBSb290IFI0NjAeFw0yMTAzMjIwMDAwMDBaFw0zNjAzMjEyMzU5
# NTlaMFQxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxKzAp
# BgNVBAMTIlNlY3RpZ28gUHVibGljIENvZGUgU2lnbmluZyBDQSBSMzYwggGiMA0G
# CSqGSIb3DQEBAQUAA4IBjwAwggGKAoIBgQCbK51T+jU/jmAGQ2rAz/V/9shTUxjI
# ztNsfvxYB5UXeWUzCxEeAEZGbEN4QMgCsJLZUKhWThj/yPqy0iSZhXkZ6Pg2A2NV
# DgFigOMYzB2OKhdqfWGVoYW3haT29PSTahYkwmMv0b/83nbeECbiMXhSOtbam+/3
# 6F09fy1tsB8je/RV0mIk8XL/tfCK6cPuYHE215wzrK0h1SWHTxPbPuYkRdkP05Zw
# mRmTnAO5/arnY83jeNzhP06ShdnRqtZlV59+8yv+KIhE5ILMqgOZYAENHNX9SJDm
# +qxp4VqpB3MV/h53yl41aHU5pledi9lCBbH9JeIkNFICiVHNkRmq4TpxtwfvjsUe
# dyz8rNyfQJy/aOs5b4s+ac7IH60B+Ja7TVM+EKv1WuTGwcLmoU3FpOFMbmPj8pz4
# 4MPZ1f9+YEQIQty/NQd/2yGgW+ufflcZ/ZE9o1M7a5Jnqf2i2/uMSWymR8r2oQBM
# dlyh2n5HirY4jKnFH/9gRvd+QOfdRrJZb1sCAwEAAaOCAWQwggFgMB8GA1UdIwQY
# MBaAFDLrkpr/NZZILyhAQnAgNpFcF4XmMB0GA1UdDgQWBBQPKssghyi47G9IritU
# pimqF6TNDDAOBgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNV
# HSUEDDAKBggrBgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEsG
# A1UdHwREMEIwQKA+oDyGOmh0dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1
# YmxpY0NvZGVTaWduaW5nUm9vdFI0Ni5jcmwwewYIKwYBBQUHAQEEbzBtMEYGCCsG
# AQUFBzAChjpodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2Rl
# U2lnbmluZ1Jvb3RSNDYucDdjMCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0
# aWdvLmNvbTANBgkqhkiG9w0BAQwFAAOCAgEABv+C4XdjNm57oRUgmxP/BP6YdURh
# w1aVcdGRP4Wh60BAscjW4HL9hcpkOTz5jUug2oeunbYAowbFC2AKK+cMcXIBD0Zd
# OaWTsyNyBBsMLHqafvIhrCymlaS98+QpoBCyKppP0OcxYEdU0hpsaqBBIZOtBajj
# cw5+w/KeFvPYfLF/ldYpmlG+vd0xqlqd099iChnyIMvY5HexjO2AmtsbpVn0OhNc
# WbWDRF/3sBp6fWXhz7DcML4iTAWS+MVXeNLj1lJziVKEoroGs9Mlizg0bUMbOalO
# hOfCipnx8CaLZeVme5yELg09Jlo8BMe80jO37PU8ejfkP9/uPak7VLwELKxAMcJs
# zkyeiaerlphwoKx1uHRzNyE6bxuSKcutisqmKL5OTunAvtONEoteSiabkPVSZ2z7
# 6mKnzAfZxCl/3dq3dUNw4rg3sTCggkHSRqTqlLMS7gjrhTqBmzu1L90Y1KWN/Y5J
# KdGvspbOrTfOXyXvmPL6E52z1NZJ6ctuMFBQZH3pwWvqURR8AgQdULUvrxjUYbHH
# j95Ejza63zdrEcxWLDX6xWls/GDnVNueKjWUH3fTv1Y8Wdho698YADR7TNx8X8z2
# Bev6SivBBOHY+uqiirZtg0y9ShQoPzmCcn63Syatatvx157YK9hlcPmVoa1oDE5/
# L9Uo2bC5a4CH2RwwggaLMIIE86ADAgECAhEAw++A/5SOEBq5lH+WfxuWEjANBgkq
# hkiG9w0BAQwFADBUMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1p
# dGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgQ0EgUjM2
# MB4XDTI2MDIyNDAwMDAwMFoXDTI3MDIyNDIzNTk1OVoweTELMAkGA1UEBhMCQVUx
# GjAYBgNVBAgMEVdlc3Rlcm4gQXVzdHJhbGlhMSYwJAYDVQQKDB1PZmZpY2Ugb2Yg
# dGhlIEF1ZGl0b3IgR2VuZXJhbDEmMCQGA1UEAwwdT2ZmaWNlIG9mIHRoZSBBdWRp
# dG9yIEdlbmVyYWwwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQD5gix/
# lAQc8KFHyrteIOU2FXQ/Vt3lJ9bH74qvOgcxN/q/rXZJ6DS9nQqf5aMbuOEP+ALs
# cZM+TnNNQEgll2x0lnCRcggfA1Odd+vmDEkNlOqX9yEvuNqJWUguVQZ6xMLqZQKX
# j93kGr16FMAb11xK7iMzhYSpWg15BBcTmvLbjjFuRhZi59l0JJ2XJAitu7sjGLrd
# cUxN39fzBnvZleWGLfvrEeRq2XTS6eL3J6yzcAo8XE9H014EaZYacVChdrfD+Vff
# XhhBjv8GeDVzgsElSw2DGW7YZ66C9YX/F3a8ZnbMXSGETv3nRQxBK+i4ov5hVdwN
# ZUc8gcA6Se03oBZSJFJCHHVDAnA0gDxuiTvUt9MEbkP2xyDmhDy+UTBhBrB5Hi7Y
# 8kn2WQxs+DRfeeEMadPGwFkhUcFFug8Eeun1Rgd3EBkLfhSg1ZFJrgjXr2zWrnAS
# wS/AoNCU9JUVLHPexsiQLb0vVnvtPU+MumjJdF2L+Yt66GipUCUhK/kgTUEpBf1L
# S2+rRPYWopkBkfdb7+RFx+d85cRDlSprYFSkrUtv5vv8Qo0ChWKq81rPGI6m85m/
# IaLoNieAdYEbiD5uCoC5C3rcnVTrn1zcL+k6j8U2pMg8ate+zxIvucrouizQ+AQ7
# HpoHSMMvsTxe5HK2zwJe90DzU1IVQLMpKDQLBwIDAQABo4IBsTCCAa0wHwYDVR0j
# BBgwFoAUDyrLIIcouOxvSK4rVKYpqhekzQwwHQYDVR0OBBYEFL9SF9gFByCoXc7Q
# TiMSt6+guomCMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQM
# MAoGCCsGAQUFBwMDMEoGA1UdIARDMEEwNQYMKwYBBAGyMQECAQMCMCUwIwYIKwYB
# BQUHAgEWF2h0dHBzOi8vc2VjdGlnby5jb20vQ1BTMAgGBmeBDAEEATBJBgNVHR8E
# QjBAMD6gPKA6hjhodHRwOi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWND
# b2RlU2lnbmluZ0NBUjM2LmNybDB5BggrBgEFBQcBAQRtMGswRAYIKwYBBQUHMAKG
# OGh0dHA6Ly9jcnQuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY0NvZGVTaWduaW5n
# Q0FSMzYuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAm
# BgNVHREEHzAdgRtTZXJ2aWNlRGVza0BhdWRpdC53YS5nb3YuYXUwDQYJKoZIhvcN
# AQEMBQADggGBAE165OjPZVPtmGXn/ds4B71BfZz2gsXKS8/85WrM5vOcp9j4QsiX
# VKPQl8dTdjVUmWk90bfW/y3zILdZXy1ZrV94rAed5ylJNgrV7/qORMK5PTQyw8Cb
# +biSLoUQmN9wCpzgy75wfqF7oAsg8xFt/QiBoSUlYXUj8SC6QLr4EU19nPjse7/s
# 97p2Th7mtzPvC0/YQHEqmlycucjwc3gmnqysMyAGWg3R+GE00ewDInaOgaMUutF5
# oEGu7KeN3r0PqmGlgp7iHUY7dEzZHg7dl1EFMxFV/awjQVwqKQ7ov530owK4pML/
# 00yx8dLms1NBEozZ7K4iVIbyB+9ZAm0pefNG0BwI2APW36MDXL+JK9KtqymCDYPm
# liBtWUwXKS6NGs6fgUVUZQrWt6bbqBMJ+r1akXsy6WqYwum3ojwW8UVN+e9V+i5z
# 6fAsMKbo0E244eB6rPfeGbGS/2sfxMgVZPV1M4Lwu7Xqzs8zR8YTQNNoi3l55EeS
# 61itZNxutwwWtjGCAxswggMXAgEBMGkwVDELMAkGA1UEBhMCR0IxGDAWBgNVBAoT
# D1NlY3RpZ28gTGltaXRlZDErMCkGA1UEAxMiU2VjdGlnbyBQdWJsaWMgQ29kZSBT
# aWduaW5nIENBIFIzNgIRAMPvgP+UjhAauZR/ln8blhIwDQYJYIZIAWUDBAIBBQCg
# gYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYB
# BAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0B
# CQQxIgQgGh/jSZ09YScRyPwjLvtPbfVH+sG4nT7c3I/DI0Y0TDAwDQYJKoZIhvcN
# AQEBBQAEggIAogNOoWHuGIVg8/uurdRpE1RsgF7KYFEHcTomva+RvTb33J6gHL2E
# mbnjObBudHNBUoTvO4wFn5em4A1DB0yxDzTH04sV5m8fSZWPCY/UBGj5AkD8qyOk
# Zp86QN7IMuWohPdFG5xM1rOIGxNl1mbrT2+dzzZb3xGtPjPmJEpmAiNMcroao7dR
# uQBgOCh/XdBd97UItJRVq6BBOuHHURl4aX4LIoQN7+mE9K6B9z6/tVjx5jAf3Ba6
# 3oUr9/dNYdhcCqPDMKQZvcCeKx9kdfHWXIPyH25kZBW4HbapdAKtiZEH6QPMzw5K
# WVmLrWcOKywx7el7RZjP3hXU2Yk5yh9/4TvGcMivl2jN3z0iw85JZz7krAbxSvjG
# ReLTcku2BYx120Vocl8VbJ/6UO5UzD+ll0DX+EbxCt7Ixsa2SZDexktFSQfJU5pG
# BhnIU2KO+nhXL5MG2FnByQxsHhGpIMZCunsyDY/G7IY6EAupsMze6DMrvRDhieSL
# SEQigSmQgMdF1P8qQDnUHO2wGI/yFiru8M1EoQunMoGssSgABdLvxDyZ2+yk3e8q
# 8pq4ukNUXHVXjEOTLUs3GY66T5O8qmSzwWQOw82GbBvIoHXchDyeKHcBqKeHzNmI
# 3TqBQuo9YwjjykOnqCLR4oRTrJqPIXGjhHSIdUBno34DE5YWfs3BOQo=
# SIG # End signature block
