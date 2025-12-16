<#
.SYNOPSIS
    Generates comprehensive Access Package reports and exports them locally.

.DESCRIPTION
    This PowerShell script connects to Microsoft Graph using certificate-based authentication to:
    
    1. Retrieve all Access Packages from Entra ID Identity Governance
    2. Analyze Access Package configurations including:
       - Resource roles and assignments
       - Approval workflows and approvers (primary, fallback, alternate across multiple stages)
       - Assignment policies and requestor settings (ALL POLICIES per Access Package)
       - Access review configurations and reviewers
       - Expiration settings and extension policies
       - Custom extensions and requestor questions
    
    3. Generate an Excel report with multiple worksheets:
       - Role Dependencies (authorization matrix)
       - AP Definitions (master list with ONE ROW PER POLICY)
       - Resource Roles breakdown
       - Reverse mapping (Role -> Access Packages)
       - Primary Approvers
       - Allowed Requesters
       - Reviewers
       - Custom Extensions
       - Requestor Questions
       - Assignments
       - Summary Statistics

.PARAMETER TenantId
    Your Azure AD Tenant ID

.PARAMETER ClientId
    The Application (client) ID of your App Registration

.PARAMETER Thumbprint
    The certificate thumbprint for authentication

.PARAMETER OutputPath
    The folder where the Excel report will be saved. Defaults to current directory.

.PARAMETER VerboseOutput
    When $True, provides detailed logging information. Default: $False

.EXAMPLE
    .\generate-access-package-report-local-v2.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -Thumbprint "your-thumbprint"

.EXAMPLE
    .\generate-access-package-report-local-v2.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -Thumbprint "your-thumbprint" -OutputPath "C:\Reports" -VerboseOutput $True

.NOTES
    Required Permissions:
    - The App Registration must have the following Microsoft Graph API permissions (Application):
        * EntitlementManagement.Read.All
        * Group.Read.All
        * Directory.Read.All
    
    Required PowerShell Modules:
        * Microsoft.Graph.Authentication
        * Microsoft.Graph.Users
        * Microsoft.Graph.Groups
        * Microsoft.Graph.Beta.Identity.Governance
        * ImportExcel
    
    Install modules with:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
        Install-Module Microsoft.Graph.Users -Scope CurrentUser
        Install-Module Microsoft.Graph.Groups -Scope CurrentUser
        Install-Module Microsoft.Graph.Beta.Identity.Governance -Scope CurrentUser
        Install-Module ImportExcel -Scope CurrentUser
    
    CHANGES IN V2:
    - Now processes ALL policies per Access Package (not just the first one)
    - Creates separate report rows for each policy
    - Better visibility into multi-policy configurations
#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$true)]
    [string]$Thumbprint,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$False)]
    [bool]$VerboseOutput = $False
)

#region Helper Functions

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [ValidateSet("Info", "Warning", "Error", "Verbose")]
        [string]$Level = "Info",
        
        [int]$Indent = 0
    )
    
    $IndentStr = "  " * $Indent
    
    switch ($Level) {
        "Info"    { Write-Host "$IndentStr$Message" -ForegroundColor Cyan }
        "Warning" { Write-Host "$IndentStr$Message" -ForegroundColor Yellow }
        "Error"   { Write-Host "$IndentStr$Message" -ForegroundColor Red }
        "Verbose" { if ($VerboseOutput) { Write-Host "$IndentStr$Message" -ForegroundColor DarkCyan } }
    }
}

function Get-ApproverDescription {
    param($Data)
    
    if ($Data.AdditionalProperties.description -ne $Null) {
        return $Data.AdditionalProperties.description
    } else {
        $Level = $Data.AdditionalProperties.managerLevel
        return "Manager($Level)"
    }
}

function Get-PolicyExtensionsAndQuestions {
    param([string]$PolicyId)
    
    $result = @{
        Extensions = @{}
        Questions = @()
    }
    
    try {
        # Single API call with all needed expansions
        $policyDetails = Get-MgBetaEntitlementManagementAccessPackageAssignmentPolicy `
            -AccessPackageAssignmentPolicyId $PolicyId `
            -ExpandProperty "CustomExtensionStageSettings(`$expand=customExtension),CustomExtensionHandlers(`$expand=customExtension)"
        
        # Process Custom Extensions - CustomExtensionStageSettings (newer)
        if ($policyDetails.CustomExtensionStageSettings) {
            foreach ($stageSetting in $policyDetails.CustomExtensionStageSettings) {
                $stage = $stageSetting.Stage
                $extensionName = $stageSetting.CustomExtension.DisplayName
                
                if ($extensionName) {
                    if (-not $result.Extensions.ContainsKey($stage)) {
                        $result.Extensions[$stage] = @()
                    }
                    $result.Extensions[$stage] += $extensionName
                }
            }
        }
        
        # Process Custom Extensions - CustomExtensionHandlers (older)
        if ($policyDetails.CustomExtensionHandlers) {
            foreach ($handler in $policyDetails.CustomExtensionHandlers) {
                $stage = $handler.Stage
                $extensionName = $handler.CustomExtension.DisplayName
                
                if ($extensionName) {
                    if (-not $result.Extensions.ContainsKey($stage)) {
                        $result.Extensions[$stage] = @()
                    }
                    $result.Extensions[$stage] += $extensionName
                }
            }
        }
        
        # Process Questions (included in the same response)
        if ($policyDetails.Questions) {
            foreach ($question in $policyDetails.Questions) {
                $questionText = $null
                if ($question.Text) {
                    if ($question.Text.DefaultText) {
                        $questionText = $question.Text.DefaultText
                    }
                    elseif ($question.Text.LocalizedTexts -and $question.Text.LocalizedTexts.Count -gt 0) {
                        $questionText = ($question.Text.LocalizedTexts | Select-Object -First 1).Text
                    }
                }
                
                $questionType = $question.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.', ''
                
                $isSingleLine = $null
                $regexPattern = $null
                $choices = @()
                $allowsMultipleSelection = $null
                
                if ($questionType -eq 'accessPackageTextInputQuestion') {
                    $isSingleLine = $question.AdditionalProperties.isSingleLineQuestion
                    $regexPattern = $question.AdditionalProperties.regexPattern
                }
                elseif ($questionType -eq 'accessPackageMultipleChoiceQuestion') {
                    $allowsMultipleSelection = $question.AdditionalProperties.allowsMultipleSelection
                    
                    if ($question.AdditionalProperties.choices) {
                        $choices = $question.AdditionalProperties.choices | ForEach-Object {
                            if ($_.displayValue.defaultText) {
                                $_.displayValue.defaultText
                            }
                            elseif ($_.displayValue.localizedTexts -and $_.displayValue.localizedTexts.Count -gt 0) {
                                ($_.displayValue.localizedTexts | Select-Object -First 1).text
                            }
                            else {
                                $_.actualValue
                            }
                        }
                    }
                }
                
                $answerFormat = switch ($questionType) {
                    'accessPackageTextInputQuestion' {
                        if ($isSingleLine -eq $true) { "Short text" }
                        elseif ($isSingleLine -eq $false) { "Long text" }
                        else { "Text" }
                    }
                    'accessPackageMultipleChoiceQuestion' {
                        if ($allowsMultipleSelection -eq $true) { "Multiple choice (multi-select)" }
                        else { "Multiple choice (single-select)" }
                    }
                    default { $questionType }
                }
                
                $result.Questions += [PSCustomObject]@{
                    Sequence             = $question.Sequence
                    QuestionText         = $questionText
                    Required             = $question.IsRequired
                    AnswerEditable       = $question.IsAnswerEditable
                    QuestionType         = $questionType
                    AnswerFormat         = $answerFormat
                    IsSingleLine         = $isSingleLine
                    RegexPattern         = $regexPattern
                    Choices              = ($choices -join "; ")
                }
            }
        }
    }
    catch {
        Write-Log -Message "Could not retrieve policy details for $PolicyId : $($_.Exception.Message)" -Level Verbose -Indent 2
    }
    
    return $result
}

function Get-ResourceRolesForEmptyPackages {
    Write-Log -Message "Building resource roles mapping for all Access Packages..."
    
    $resourceRolesForEmptyAccessPackages = @{}
    
    try {
        $AccessPackageCatalogObjList = Get-MgBetaEntitlementManagementAccessPackageCatalog -All
        $CatalogCount = $AccessPackageCatalogObjList.Count
        $CatalogIndex = 0
        
        Write-Log -Message "Retrieved $CatalogCount catalogs" -Level Verbose -Indent 1
        
        foreach ($AccessPackageCatalogObj in $AccessPackageCatalogObjList) {
            $CatalogIndex++
            $CatalogPercent = [math]::Round(($CatalogIndex / $CatalogCount) * 100)
            
            # Progress for catalog processing
            Write-Progress -Activity "Building Resource Roles Mapping" `
                -Status "[$CatalogIndex/$CatalogCount] Catalog: $($AccessPackageCatalogObj.DisplayName)" `
                -PercentComplete $CatalogPercent
            
            Write-Log -Message "Processing catalog: $($AccessPackageCatalogObj.DisplayName)" -Level Verbose -Indent 1
            
            try {
                $AccessPackageCatalogAccessPackageResourceList = Get-MgBetaEntitlementManagementAccessPackageCatalogAccessPackageResource `
                    -AccessPackageCatalogId $AccessPackageCatalogObj.Id `
                    -ExpandProperty "AccessPackageResourceRoles,AccessPackageResourceScopes,AccessPackageResourceEnvironment"
            }
            catch {
                Write-Log -Message "Failed to get resources for catalog '$($AccessPackageCatalogObj.DisplayName)': $($_.Exception.Message)" -Level Warning -Indent 2
                continue
            }

            try {
                $AccessPackageObjList = Get-MgBetaEntitlementManagementAccessPackage -CatalogId $AccessPackageCatalogObj.Id -ExpandProperty AccessPackageResourceRoleScopes
            }
            catch {
                Write-Log -Message "Failed to get Access Packages for catalog '$($AccessPackageCatalogObj.DisplayName)': $($_.Exception.Message)" -Level Warning -Indent 2
                continue
            }

            foreach ($AccessPackageObj in $AccessPackageObjList) {
                try {
                    $AccessPackageRoleAndScopeObjList = $AccessPackageObj.AccessPackageResourceRoleScopes | 
                        Select-Object -ExpandProperty Id | 
                        Select-Object @{E={$_ -split("_") | Select-Object -First 1};L="RoleId"},
                                      @{E={$_ -split("_") | Select-Object -Last 1};L="ScopeId"}
                    
                    [Array]$RSRoleArray = @()
                    foreach ($AccessPackageRoleAndScopeObj in $AccessPackageRoleAndScopeObjList) {
                        [Array]$RoleObjList = $AccessPackageCatalogAccessPackageResourceList | 
                            Where-Object {$_.AccessPackageResourceRoles.id -contains $AccessPackageRoleAndScopeObj.RoleId}
                        
                        foreach ($RoleObj in $RoleObjList) {
                            $RSRoleArray += $RoleObj.DisplayName
                        }
                    }
                    
                    $resourceRolesForEmptyAccessPackages[$AccessPackageObj.DisplayName] = $RSRoleArray
                }
                catch {
                    Write-Log -Message "Failed to process AccessPackage '$($AccessPackageObj.DisplayName)': $($_.Exception.Message)" -Level Warning -Indent 2
                }
            }
        }
        
        Write-Progress -Activity "Building Resource Roles Mapping" -Completed
    }
    catch {
        Write-Log -Message "Error building resource roles mapping: $($_.Exception.Message)" -Level Error
    }
    
    Write-Log -Message "Completed resource roles mapping for $($resourceRolesForEmptyAccessPackages.Keys.Count) Access Packages" -Indent 1
    return $resourceRolesForEmptyAccessPackages
}

#endregion

#region Main Script

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

Write-Host ""
Write-Host "=============================================" -ForegroundColor White
Write-Host "   Access Package Report Generator v2        " -ForegroundColor White
Write-Host "   (Multi-Policy Support)                    " -ForegroundColor White
Write-Host "=============================================" -ForegroundColor White
Write-Host ""

# Validate output path
if (!(Test-Path -Path $OutputPath -PathType Container)) {
    Write-Log -Message "Creating output directory: $OutputPath" -Level Info
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# Get certificate
Write-Log -Message "Looking for certificate with thumbprint: $Thumbprint"
$CertificateObject = Get-ChildItem "Cert:\CurrentUser\My" | 
    Where-Object { $_.Thumbprint -eq $Thumbprint } | 
    Sort-Object -Descending NotAfter | 
    Select-Object -First 1

if ($null -eq $CertificateObject) {
    Write-Log -Message "Certificate not found in CurrentUser store. Checking LocalMachine store..." -Level Warning
    $CertificateObject = Get-ChildItem "Cert:\LocalMachine\My" | 
        Where-Object { $_.Thumbprint -eq $Thumbprint } | 
        Sort-Object -Descending NotAfter | 
        Select-Object -First 1
}

if ($null -eq $CertificateObject) {
    Write-Log -Message "Certificate with thumbprint '$Thumbprint' not found!" -Level Error
    exit 1
}

Write-Log -Message "Certificate found: $($CertificateObject.Subject)" -Indent 1
Write-Log -Message "Expires: $($CertificateObject.NotAfter)" -Indent 1

# Check certificate expiration
$DaysUntilExpiry = ($CertificateObject.NotAfter - (Get-Date)).Days
if ($DaysUntilExpiry -lt 0) {
    Write-Log -Message "Certificate has expired!" -Level Error
    exit 1
}
elseif ($DaysUntilExpiry -lt 30) {
    Write-Log -Message "Warning: Certificate expires in $DaysUntilExpiry days!" -Level Warning
}

# Connect to Microsoft Graph
Write-Log -Message "Connecting to Microsoft Graph..."
try {
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome
    Write-Log -Message "Connected successfully" -Level Info -Indent 1
}
catch {
    Write-Log -Message "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level Error
    exit 1
}

# Build resource roles mapping
$resourcesEmptyAccessPackagesAssignments = Get-ResourceRolesForEmptyPackages

# Get Access Packages
Write-Log -Message "Retrieving Access Packages..."
$AccessPackageList = Get-MgBetaEntitlementManagementAccessPackage -All -ExpandProperty AccessPackageCatalog
$AccessPackageList = $AccessPackageList | Sort-Object @{e={$_.AccessPackageCatalog.DisplayName}}, DisplayName
Write-Log -Message "Found $($AccessPackageList.Count) Access Packages" -Indent 1

# Get Entitlement Management Access Package list with policies
Write-Log -Message "Retrieving Access Package policies..."
$EntitlementManagementAccessPackageList = Get-MgBetaEntitlementManagementAccessPackage -ExpandProperty AccessPackageAssignmentPolicies -All

# Initialize report collections
$Report = @()
$ReportQuestions = @()
$ReportCustomExtensions = @()
$ReportAssignments = [System.Collections.Generic.List[PSObject]]::new()

$CountTotal = $AccessPackageList.Count
$CurrentIndex = 0
$StartTime = Get-Date
$TotalAssignmentsFound = 0
$TotalPoliciesProcessed = 0

Write-Log -Message "Processing Access Packages..."
Write-Host ""

foreach ($AccessPackage in $AccessPackageList) {
    $CurrentIndex++
    $PercentComplete = [math]::Round(($CurrentIndex / $CountTotal) * 100)
    
    $AccessPackageName = $AccessPackage.DisplayName
    $AccessPackageID = $AccessPackage.Id
    $Catalog = $AccessPackage.AccessPackageCatalog.DisplayName
    $Description = $AccessPackage.Description

    # Calculate elapsed time and estimate remaining time
    $ElapsedTime = (Get-Date) - $StartTime
    $AvgTimePerPackage = if ($CurrentIndex -gt 1) { $ElapsedTime.TotalSeconds / ($CurrentIndex - 1) } else { 0 }
    $EstimatedRemaining = [TimeSpan]::FromSeconds($AvgTimePerPackage * ($CountTotal - $CurrentIndex))
    
    # Build progress status
    $ProgressStatus = "[$CurrentIndex/$CountTotal] $Catalog / $AccessPackageName"
    
    # Write progress bar
    Write-Progress -Activity "Processing Access Packages" `
        -Status $ProgressStatus `
        -PercentComplete $PercentComplete `
        -SecondsRemaining $EstimatedRemaining.TotalSeconds
    
    # Write console progress (overwrites same line)
    $ProgressBar = ""
    $BarLength = 30
    $FilledLength = [math]::Round($BarLength * $PercentComplete / 100)
    $ProgressBar = ("█" * $FilledLength) + ("░" * ($BarLength - $FilledLength))
    
    $StatusLine = "`r  [$ProgressBar] $PercentComplete% | $CurrentIndex/$CountTotal | ETA: $($EstimatedRemaining.ToString('hh\:mm\:ss')) | $AccessPackageName"
    # Truncate if too long and pad to overwrite previous text
    if ($StatusLine.Length -gt 120) { $StatusLine = $StatusLine.Substring(0, 117) + "..." }
    $StatusLine = $StatusLine.PadRight(130)
    Write-Host $StatusLine -NoNewline -ForegroundColor Gray
    
    Write-Log -Message "[$CurrentIndex/$CountTotal] $AccessPackageName" -Level Verbose -Indent 1
    
    # Get resource roles from catalog mapping (works for all packages, regardless of assignments)
    $GroupsString = $resourcesEmptyAccessPackagesAssignments[$AccessPackageName] -join " | "
    
    # Get ALL assignments for the assignments report (ONCE per Access Package)
    $AllAssignments = Get-MgBetaEntitlementManagementAccessPackageAssignment `
        -AccessPackageId $AccessPackageID `
        -ExpandProperty target -All
    
    $TotalAssignmentsFound += $AllAssignments.Count
    
    if ($AllAssignments.Count -eq 0) {
        $ReportAssignments.Add([PSCustomObject]@{
            "AP Name" = $AccessPackageName
            "Catalog" = $Catalog
            "Assignment" = ""
            "UPN" = ""
            "Assignment State" = ""
        })
    } else {
        foreach ($Assignment in $AllAssignments) {
            $ReportAssignments.Add([PSCustomObject]@{
                "AP Name" = $AccessPackageName
                "Catalog" = $Catalog
                "Assignment" = $Assignment.Target.DisplayName
                "UPN" = $Assignment.Target.PrincipalName
                "Assignment State" = $Assignment.AssignmentState
            })
        }
    }

    # Get ALL Assignment Policies for this Access Package
    [Array]$AccessPackageAssignmentPolicies = $EntitlementManagementAccessPackageList | 
        Where-Object { $_.Id -eq $AccessPackageID } | 
        Select-Object -ExpandProperty AccessPackageAssignmentPolicies
    
    Write-Log -Message "  Found $($AccessPackageAssignmentPolicies.Count) policies for this Access Package" -Level Verbose -Indent 2
    
    # Get Custom Extensions and Questions (ONCE per Access Package, from ALL policies)
    $AllCustomExtensions = @{}
    $QuestionsFromFirstPolicy = @()
    
    foreach ($Policy in $AccessPackageAssignmentPolicies) {
        $PolicyDetails = Get-PolicyExtensionsAndQuestions -PolicyId $Policy.Id
        
        # Collect extensions from all policies
        foreach ($Stage in $PolicyDetails.Extensions.Keys) {
            if (-not $AllCustomExtensions.ContainsKey($Stage)) {
                $AllCustomExtensions[$Stage] = @()
            }
            $AllCustomExtensions[$Stage] += $PolicyDetails.Extensions[$Stage]
        }
        
        # Get questions from first policy only (questions are typically the same across policies)
        if ($Policy.Id -eq $AccessPackageAssignmentPolicies[0].Id) {
            $QuestionsFromFirstPolicy = $PolicyDetails.Questions
        }
    }
    
    # Build Custom Extensions string (once per AP)
    $CustomExtensionsString = ""
    if ($AllCustomExtensions.Count -gt 0) {
        $ExtParts = @()
        foreach ($Stage in ($AllCustomExtensions.Keys | Sort-Object)) {
            $ExtNames = ($AllCustomExtensions[$Stage] | Select-Object -Unique) -join ", "
            $ExtParts += "$Stage`: $ExtNames"
        }
        $CustomExtensionsString = $ExtParts -join " | "
    }
    
    # Add to Custom Extensions report (once per AP)
    foreach ($Stage in $AllCustomExtensions.Keys) {
        foreach ($ExtName in ($AllCustomExtensions[$Stage] | Select-Object -Unique)) {
            $ReportCustomExtensions += [PSCustomObject][Ordered]@{
                "AP Name" = $AccessPackageName
                "AP ID" = $AccessPackageID
                "Catalog" = $Catalog
                "Extension Stage" = $Stage
                "Extension Name" = $ExtName
            }
        }
    }
    
    # Process Questions (once per AP, from first policy)
    $QuestionsCount = $QuestionsFromFirstPolicy.Count
    $QuestionsString = ""
    
    if ($QuestionsFromFirstPolicy.Count -gt 0) {
        $QuestionTexts = $QuestionsFromFirstPolicy | ForEach-Object { 
            $reqMark = if ($_.Required) { "*" } else { "" }
            "$($_.QuestionText)$reqMark"
        }
        $QuestionsString = $QuestionTexts -join " | "
        
        foreach ($Q in $QuestionsFromFirstPolicy) {
            $ReportQuestions += [PSCustomObject][Ordered]@{
                "AP Name" = $AccessPackageName
                "AP ID" = $AccessPackageID
                "Catalog" = $Catalog
                "Policy Name" = $AccessPackageAssignmentPolicies[0].DisplayName
                "Sequence" = $Q.Sequence
                "Question Text" = $Q.QuestionText
                "Required" = $Q.Required
                "Answer Format" = $Q.AnswerFormat
                "Is Single Line" = $Q.IsSingleLine
                "Regex Pattern" = $Q.RegexPattern
                "Choices" = $Q.Choices
                "Answer Editable" = $Q.AnswerEditable
            }
        }
    }

    # NOW LOOP THROUGH ALL POLICIES to create separate report rows
    foreach ($Policy in $AccessPackageAssignmentPolicies) {
        $TotalPoliciesProcessed++
        
        $AssignmentPolicyDisplayName = $Policy.DisplayName
        $ApprovalMode = $Policy.RequestApprovalSettings.ApprovalMode
        
        Write-Log -Message "    Processing policy: $AssignmentPolicyDisplayName" -Level Verbose -Indent 3
        
        # Allowed requestors
        $AllowedRequestors = $Null
        $Policy.RequestorSettings.AllowedRequestors | ForEach-Object {
            $Req = $_.AdditionalProperties.description
            $AllowedRequestors += "$Req |"
        }

        # Get approvers for all stages
        $PrimaryApproversStage1 = $Null
        $FallbackApproversStage1 = $Null
        $PrimaryApproversStage2 = $Null
        $FallbackApproversStage2 = $Null
        $PrimaryApproversStage3 = $Null
        $FallbackApproversStage3 = $Null

        if ($ApprovalMode -ne "NoApproval") {
            $counter = ($Policy.RequestApprovalSettings.ApprovalStages).Count
            for ($x = 0; $x -lt $counter; $x++) {
                $Stage = $x + 1
                $Policy.RequestApprovalSettings.ApprovalStages[$x].PrimaryApprovers | ForEach-Object {
                    $Approver = Get-ApproverDescription -Data $_
                    $IsBackup = $_.IsBackup
                    switch ($Stage) {
                        1 { if ($IsBackup) { $FallbackApproversStage1 += "$Approver | " } else { $PrimaryApproversStage1 += "$Approver | " } }
                        2 { if ($IsBackup) { $FallbackApproversStage2 += "$Approver | " } else { $PrimaryApproversStage2 += "$Approver | " } }
                        3 { if ($IsBackup) { $FallbackApproversStage3 += "$Approver | " } else { $PrimaryApproversStage3 += "$Approver | " } }
                    }
                }
            }
        }
        
        # Get review settings
        $ReviewRecurrence = $Policy.AccessReviewSettings.RecurrenceType
        
        $Reviewers = $Null
        if ($Policy.AccessReviewSettings.ReviewerType -eq 'Manager') {
            $Reviewers = "Manager"
        } else {
            $Reviewers = (($Policy.AccessReviewSettings.Reviewers | ForEach-Object { $_.AdditionalProperties.description }) -join " | ")
        }
        
        # Request settings
        $EnableNewRequests = $Policy.RequestorSettings.AcceptRequests

        # Expiration settings
        $Expiration = $Policy.DurationInDays
        $canExtend = $Policy.CanExtend
        $IsApprovalRequiredForExtension = $Policy.RequestApprovalSettings.IsApprovalRequiredForExtension

        # Alternate/Escalation approvers
        $AlternateApprover = $Null
        $AlternateApproverFallback = $Null
        $AlternateApproverSecondLevel = $Null
        $AlternateApproverFallbackSecondLevel = $Null
        $AlternateApproverThirdLevel = $Null
        $AlternateApproverFallbackThirdLevel = $Null
        
        if ($Policy.RequestApprovalSettings.ApprovalStages.Count -ge 1 -and $Policy.RequestApprovalSettings.ApprovalStages[0].IsEscalationEnabled) {
            foreach ($Item in $Policy.RequestApprovalSettings.ApprovalStages[0].EscalationApprovers) {
                $Approver = Get-ApproverDescription -Data $Item
                if ($Item.IsBackup -eq $true) {
                    $AlternateApproverFallback += "$Approver | "
                } else {
                    $AlternateApprover += "$Approver | "
                }
            }
        }
        
        if ($Policy.RequestApprovalSettings.ApprovalStages.Count -ge 2 -and $Policy.RequestApprovalSettings.ApprovalStages[1].IsEscalationEnabled) {
            foreach ($Item in $Policy.RequestApprovalSettings.ApprovalStages[1].EscalationApprovers) {
                $Approver = Get-ApproverDescription -Data $Item
                if ($Item.IsBackup -eq $true) {
                    $AlternateApproverFallbackSecondLevel += "$Approver | "
                } else {
                    $AlternateApproverSecondLevel += "$Approver | "
                }
            }
        }
        
        if ($Policy.RequestApprovalSettings.ApprovalStages.Count -ge 3 -and $Policy.RequestApprovalSettings.ApprovalStages[2].IsEscalationEnabled) {
            foreach ($Item in $Policy.RequestApprovalSettings.ApprovalStages[2].EscalationApprovers) {
                $Approver = Get-ApproverDescription -Data $Item
                if ($Item.IsBackup -eq $true) {
                    $AlternateApproverFallbackThirdLevel += "$Approver | "
                } else {
                    $AlternateApproverThirdLevel += "$Approver | "
                }
            }
        }

        # Build main report object (ONE ROW PER POLICY)
        $Obj = [PSCustomObject][Ordered]@{
            "AP Name" = $AccessPackageName
            "AP Description" = $Description
            "AP ID" = $AccessPackageID
            "Catalog" = $Catalog
            "Resource Roles" = $GroupsString
            "Policy ID" = $Policy.Id
            "Policy Display Name" = $AssignmentPolicyDisplayName
            "Approval Mode" = $ApprovalMode
            "Allowed Requesters" = $AllowedRequestors
            "Primary Approvers" = $PrimaryApproversStage1
            "Fallback Approvers" = $FallbackApproversStage1
            "Alternate Approvers" = $AlternateApprover
            "Alternate Approvers Fallback" = $AlternateApproverFallback
            "2nd Stage Approvers" = $PrimaryApproversStage2
            "2nd Stage Fallback Approvers" = $FallbackApproversStage2
            "2nd Stage Alternate Approvers" = $AlternateApproverSecondLevel
            "2nd Stage Alternate Approvers Fallback" = $AlternateApproverFallbackSecondLevel
            "3rd Stage Approvers" = $PrimaryApproversStage3
            "3rd Stage Fallback Approvers" = $FallbackApproversStage3
            "3rd Stage Alternate Approvers" = $AlternateApproverThirdLevel
            "3rd Stage Alternate Approvers Fallback" = $AlternateApproverFallbackThirdLevel
            "Review Recurrence" = $ReviewRecurrence
            "Reviewers" = $Reviewers
            "New Requests Enabled" = $EnableNewRequests
            "Expiration in Days" = $Expiration
            "Can Extend" = $canExtend
            "Extension Approval Required" = $IsApprovalRequiredForExtension
            "Custom Extensions" = $CustomExtensionsString
            "Questions Count" = $QuestionsCount
            "Questions" = $QuestionsString
        }

        $Report += $Obj
    }
}

Write-Progress -Activity "Processing Access Packages" -Completed

# Clear the progress line and show completion
$TotalElapsedTime = (Get-Date) - $StartTime
Write-Host "`r".PadRight(135) # Clear the line
Write-Host ""
Write-Host "  Processing complete!" -ForegroundColor Green
Write-Host "  ├─ Access Packages processed: $CountTotal" -ForegroundColor Gray
Write-Host "  ├─ Total policies processed: $TotalPoliciesProcessed" -ForegroundColor Gray
Write-Host "  ├─ Total assignments found: $TotalAssignmentsFound" -ForegroundColor Gray
Write-Host "  └─ Time elapsed: $($TotalElapsedTime.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
Write-Host ""

# Build additional report views
Write-Log -Message "Building additional report views..."

# Resource Roles report
$ReportResourceRoles = @()
$Hashmap = [hashtable]::new()
$AllowedRequesters = @()
$ReportReviewers = @()
$ReportPrimaryApprovers = @()

foreach ($Item in $Report) {
    $APName = $Item."AP Name"
    $PolicyName = $Item."Policy Display Name"
    $ResourceRoles = $Item."Resource Roles"
    $PrimApprovers = $Item."Primary Approvers"
    $Requesters = $Item."Allowed Requesters"
    $Reviewers = $Item."Reviewers"

    # Resource Roles (one entry per policy)
    if ([string]::IsNullOrEmpty($ResourceRoles)) {
        $ReportResourceRoles += [PSCustomObject]@{ 
            "AP Name" = $APName
            "Policy Name" = $PolicyName
            "Resource Role" = "" 
        }
    } else {
        foreach ($x in ($ResourceRoles.Split('|') | Where-Object { ![string]::IsNullOrWhiteSpace($_) })) {
            $x = $x.Trim()
            $ReportResourceRoles += [PSCustomObject]@{ 
                "AP Name" = $APName
                "Policy Name" = $PolicyName
                "Resource Role" = $x 
            }
            # For reverse mapping, use AP name only (not policy-specific)
            if ($Hashmap.ContainsKey($x)) {
                if ($Hashmap[$x] -notlike "*$APName*") {
                    $Hashmap[$x] += "| $APName"
                }
            } else {
                $Hashmap.Add($x, "$APName")
            }
        }
    }

    # Primary Approvers (one entry per policy)
    if ([string]::IsNullOrEmpty($PrimApprovers)) {
        $ReportPrimaryApprovers += [PSCustomObject]@{ 
            "AP Name" = $APName
            "Policy Name" = $PolicyName
            "Primary Approver" = "" 
        }
    } else {
        foreach ($x in ($PrimApprovers.Split('|') | Where-Object { ![string]::IsNullOrWhiteSpace($_) })) {
            $ReportPrimaryApprovers += [PSCustomObject]@{ 
                "AP Name" = $APName
                "Policy Name" = $PolicyName
                "Primary Approver" = $x.Trim() 
            }
        }
    }

    # Allowed Requesters (one entry per policy)
    if ([string]::IsNullOrEmpty($Requesters)) {
        $AllowedRequesters += [PSCustomObject]@{ 
            "AP Name" = $APName
            "Policy Name" = $PolicyName
            "Allowed Requesters" = "" 
        }
    } else {
        foreach ($x in ($Requesters.Split('|') | Where-Object { ![string]::IsNullOrWhiteSpace($_) })) {
            $AllowedRequesters += [PSCustomObject]@{ 
                "AP Name" = $APName
                "Policy Name" = $PolicyName
                "Allowed Requesters" = $x.Trim() 
            }
        }
    }

    # Reviewers (one entry per policy)
    if ([string]::IsNullOrEmpty($Reviewers)) {
        $ReportReviewers += [PSCustomObject]@{ 
            "AP Name" = $APName
            "Policy Name" = $PolicyName
            "Reviewers" = "" 
        }
    } else {
        foreach ($x in ($Reviewers.Split('|') | Where-Object { ![string]::IsNullOrWhiteSpace($_) })) {
            $ReportReviewers += [PSCustomObject]@{ 
                "AP Name" = $APName
                "Policy Name" = $PolicyName
                "Reviewers" = $x.Trim() 
            }
        }
    }
}

# Hashmap decoupling (reverse mapping: Role -> Access Packages)
$HashmapDecoupling = @()
foreach ($Entry in $Hashmap.GetEnumerator()) {
    if ([string]::IsNullOrEmpty($Entry.Value)) {
        $HashmapDecoupling += [PSCustomObject]@{ "Role" = $Entry.Key; "Access Package" = "" }
    } else {
        foreach ($x in ($Entry.Value.Split('|') | Where-Object { ![string]::IsNullOrWhiteSpace($_) })) {
            $HashmapDecoupling += [PSCustomObject]@{ "Role" = $Entry.Key; "Access Package" = $x.Trim() }
        }
    }
}

# Build Role Dependencies report (OPTIMIZED)
Write-Log -Message "Building Role Dependencies report (Optimized)..."

$ReportSummary = [System.Collections.Generic.List[PSObject]]::new()
$DependencyIndex = 0
$DependencyTotal = $Report.Count

foreach ($R in $Report) {
    $DependencyIndex++
    if ($DependencyIndex % 50 -eq 0) {
        Write-Progress -Activity "Building Role Dependencies" -Status "Processing policy $DependencyIndex of $DependencyTotal" -PercentComplete (($DependencyIndex / $DependencyTotal) * 100)
    }

    # Temporary Hashtable for this package/policy to track unique actors
    # Key = User/Group Name, Value = Hashtable of their roles in this package
    $PackageActors = @{} 

    # Helper scriptblock to process a column and add actors to the list
    $ProcessColumn = {
        param($ColName, $TargetField)
        if (-not [string]::IsNullOrWhiteSpace($R.$ColName)) {
            $Parts = $R.$ColName -split '\|'
            foreach ($P in $Parts) {
                $P = $P.Trim()
                if (-not [string]::IsNullOrWhiteSpace($P)) {
                    if (-not $PackageActors.ContainsKey($P)) { $PackageActors[$P] = @{} }
                    $PackageActors[$P][$TargetField] = "X"
                }
            }
        }
    }

    # Map the Report columns to the Output columns
    &$ProcessColumn "Allowed Requesters" "Allowed Requester"
    &$ProcessColumn "Primary Approvers" "Primary Approver"
    &$ProcessColumn "Fallback Approvers" "Fallback Approver"
    &$ProcessColumn "Alternate Approvers" "Alternate Approvers"
    &$ProcessColumn "Alternate Approvers Fallback" "Alternate Approvers Fallback"
    &$ProcessColumn "2nd Stage Approvers" "2nd Stage Approver"
    &$ProcessColumn "2nd Stage Fallback Approvers" "2nd Stage Fallback Approver"
    &$ProcessColumn "2nd Stage Alternate Approvers" "2nd Stage Alternate Approvers"
    &$ProcessColumn "2nd Stage Alternate Approvers Fallback" "2nd Stage Alternate Approvers Fallback"
    &$ProcessColumn "3rd Stage Approvers" "3rd Stage Approver"
    &$ProcessColumn "3rd Stage Fallback Approvers" "3rd Stage Fallback Approver"
    &$ProcessColumn "3rd Stage Alternate Approvers" "3rd Stage Alternate Approvers"
    &$ProcessColumn "3rd Stage Alternate Approvers Fallback" "3rd Stage Alternate Approvers Fallback"
    &$ProcessColumn "Reviewers" "Reviewer"
    &$ProcessColumn "Resource Roles" "Resource Role"

    # Convert the found actors into report rows for this package/policy
    foreach ($ActorName in $PackageActors.Keys) {
        $Roles = $PackageActors[$ActorName]
        
        $ReportSummary.Add([PSCustomObject][Ordered]@{
            "Access Package" = $R.'AP Name'
            "Policy Name" = $R.'Policy Display Name'
            "Access Package Description" = $R.'AP Description'
            "User/Group" = $ActorName
            "Allowed Requester" = if ($Roles.ContainsKey("Allowed Requester")) { "X" } else { "" }
            "Primary Approver" = if ($Roles.ContainsKey("Primary Approver")) { "X" } else { "" }
            "Fallback Approver" = if ($Roles.ContainsKey("Fallback Approver")) { "X" } else { "" }
            "Alternate Approvers" = if ($Roles.ContainsKey("Alternate Approvers")) { "X" } else { "" }
            "Alternate Approvers Fallback" = if ($Roles.ContainsKey("Alternate Approvers Fallback")) { "X" } else { "" }
            "2nd Stage Approver" = if ($Roles.ContainsKey("2nd Stage Approver")) { "X" } else { "" }
            "2nd Stage Fallback Approver" = if ($Roles.ContainsKey("2nd Stage Fallback Approver")) { "X" } else { "" }
            "2nd Stage Alternate Approvers" = if ($Roles.ContainsKey("2nd Stage Alternate Approvers")) { "X" } else { "" }
            "2nd Stage Alternate Approvers Fallback" = if ($Roles.ContainsKey("2nd Stage Alternate Approvers Fallback")) { "X" } else { "" }
            "3rd Stage Approver" = if ($Roles.ContainsKey("3rd Stage Approver")) { "X" } else { "" }
            "3rd Stage Fallback Approver" = if ($Roles.ContainsKey("3rd Stage Fallback Approver")) { "X" } else { "" }
            "3rd Stage Alternate Approvers" = if ($Roles.ContainsKey("3rd Stage Alternate Approvers")) { "X" } else { "" }
            "3rd Stage Alternate Approvers Fallback" = if ($Roles.ContainsKey("3rd Stage Alternate Approvers Fallback")) { "X" } else { "" }
            "Reviewer" = if ($Roles.ContainsKey("Reviewer")) { "X" } else { "" }
            "Resource Role" = if ($Roles.ContainsKey("Resource Role")) { "X" } else { "" }
        })
    }
}

Write-Progress -Activity "Building Role Dependencies" -Completed
Write-Host "  Role Dependencies: $($ReportSummary.Count) entries created" -ForegroundColor Gray

# Summary Statistics
$UniqueAccessPackages = ($Report | Select-Object -Property "AP Name" -Unique).Count
$SummaryStats = @(
    [PSCustomObject]@{ Metric = "Total Access Packages"; Value = $UniqueAccessPackages }
    [PSCustomObject]@{ Metric = "Total Assignment Policies"; Value = $Report.Count }
    [PSCustomObject]@{ Metric = "Total Catalogs"; Value = ($Report | Select-Object -ExpandProperty Catalog -Unique).Count }
    [PSCustomObject]@{ Metric = "Access Packages with Multiple Policies"; Value = ($Report | Group-Object "AP Name" | Where-Object { $_.Count -gt 1 }).Count }
    [PSCustomObject]@{ Metric = "Policies with Approval Required"; Value = ($Report | Where-Object { $_."Approval Mode" -ne "NoApproval" }).Count }
    [PSCustomObject]@{ Metric = "Policies with No Approval"; Value = ($Report | Where-Object { $_."Approval Mode" -eq "NoApproval" }).Count }
    [PSCustomObject]@{ Metric = "Policies with Access Reviews"; Value = ($Report | Where-Object { $_."Review Recurrence" -ne $null -and $_."Review Recurrence" -ne "" }).Count }
    [PSCustomObject]@{ Metric = "Access Packages with Custom Extensions"; Value = ($ReportCustomExtensions | Select-Object "AP Name" -Unique).Count }
    [PSCustomObject]@{ Metric = "Access Packages with Requestor Questions"; Value = ($ReportQuestions | Select-Object "AP Name" -Unique).Count }
    [PSCustomObject]@{ Metric = "Total Custom Extension Configurations"; Value = $ReportCustomExtensions.Count }
    [PSCustomObject]@{ Metric = "Total Requestor Questions"; Value = $ReportQuestions.Count }
    [PSCustomObject]@{ Metric = "Total Assignments"; Value = ($ReportAssignments | Where-Object { $_."Assignment" -ne "" }).Count }
    [PSCustomObject]@{ Metric = "Report Generated"; Value = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
)

# Export to Excel
Write-Log -Message "Exporting report to Excel..."

$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
$TableStyle = "Medium2"

# Main comprehensive report
$ReportPath = Join-Path $OutputPath "AccessPackageReport_$Timestamp.xlsx"
Write-Log -Message "Creating: $ReportPath" -Indent 1

$ReportSummary | Export-Excel -Path $ReportPath -WorksheetName "Role_Dependencies" -TableStyle $TableStyle -TableName "Role_Dependencies" -AutoSize -FreezeTopRow -BoldTopRow
$Report | Export-Excel -Path $ReportPath -WorksheetName "AP_Definitions" -TableStyle $TableStyle -TableName "AP_Definitions" -AutoSize -FreezeTopRow -BoldTopRow
$ReportResourceRoles | Export-Excel -Path $ReportPath -WorksheetName "AP_ResourceRoles" -TableStyle $TableStyle -TableName "AP_ResourceRoles" -AutoSize -FreezeTopRow -BoldTopRow
$HashmapDecoupling | Export-Excel -Path $ReportPath -WorksheetName "ResourceRoles_AP" -TableStyle $TableStyle -TableName "ResourceRoles_AP" -AutoSize -FreezeTopRow -BoldTopRow
$ReportPrimaryApprovers | Export-Excel -Path $ReportPath -WorksheetName "AP_PrimaryApprovers" -TableStyle $TableStyle -TableName "AP_PrimaryApprovers" -AutoSize -FreezeTopRow -BoldTopRow
$AllowedRequesters | Export-Excel -Path $ReportPath -WorksheetName "AP_AllowedRequesters" -TableStyle $TableStyle -TableName "AP_AllowedRequesters" -AutoSize -FreezeTopRow -BoldTopRow
$ReportReviewers | Export-Excel -Path $ReportPath -WorksheetName "AP_Reviewers" -TableStyle $TableStyle -TableName "AP_Reviewers" -AutoSize -FreezeTopRow -BoldTopRow
$ReportAssignments | Export-Excel -Path $ReportPath -WorksheetName "AP_Assignments" -TableStyle $TableStyle -TableName "AP_Assignments" -AutoSize -FreezeTopRow -BoldTopRow

if ($ReportCustomExtensions.Count -gt 0) {
    $ReportCustomExtensions | Export-Excel -Path $ReportPath -WorksheetName "AP_CustomExtensions" -TableStyle $TableStyle -TableName "AP_CustomExtensions" -AutoSize -FreezeTopRow -BoldTopRow
}

if ($ReportQuestions.Count -gt 0) {
    $ReportQuestions | Export-Excel -Path $ReportPath -WorksheetName "AP_Questions" -TableStyle $TableStyle -TableName "AP_Questions" -AutoSize -FreezeTopRow -BoldTopRow
}

$SummaryStats | Export-Excel -Path $ReportPath -WorksheetName "Summary" -TableStyle $TableStyle -TableName "Summary" -AutoSize -FreezeTopRow -BoldTopRow

# Disconnect
Write-Log -Message "Disconnecting from Microsoft Graph..."
Disconnect-MgGraph | Out-Null

# Summary
Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "   Report Generation Complete!" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "File created:" -ForegroundColor White
Write-Host "  $ReportPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "Summary:" -ForegroundColor White
Write-Host "  - Unique Access Packages: $UniqueAccessPackages" -ForegroundColor Cyan
Write-Host "  - Total Assignment Policies: $($Report.Count)" -ForegroundColor Cyan
Write-Host "  - Access Packages with Multiple Policies: $(($Report | Group-Object 'AP Name' | Where-Object { $_.Count -gt 1 }).Count)" -ForegroundColor Cyan
Write-Host "  - Total Catalogs: $(($Report | Select-Object -ExpandProperty Catalog -Unique).Count)" -ForegroundColor Cyan
Write-Host "  - Total Assignments: $(($ReportAssignments | Where-Object { $_.'Assignment' -ne '' }).Count)" -ForegroundColor Cyan
Write-Host "  - Role Dependencies entries: $($ReportSummary.Count)" -ForegroundColor Cyan
Write-Host ""

#endregion
