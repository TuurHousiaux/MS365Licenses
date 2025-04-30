#Requires -Version 7.0

# Import only the required Microsoft Graph modules
Write-Progress -Activity "Loading Microsoft Graph Modules" -Status "Please wait..."
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Users.Actions -ErrorAction Stop

# License SKU to friendly name mapping
$licenseMapping = @{
    'Microsoft_Teams_EEA_New' = 'Microsoft Teams (EEA)'
    'O365_w/o_Teams_Bundle_M5' = 'Office 365 E5 (without Teams)'
    'ENTERPRISEPACK' = 'Office 365 E3'
    'ENTERPRISEPREMIUM' = 'Office 365 E5'
    'STANDARDPACK' = 'Office 365 E1'
    'EXCHANGESTANDARD' = 'Exchange Online Plan 1'
    'EXCHANGEENTERPRISE' = 'Exchange Online Plan 2'
    'SPE_E3' = 'Microsoft 365 E3'
    'SPE_E5' = 'Microsoft 365 E5'
    'MCOPSTN2' = 'Microsoft 365 Domestic Calling Plan'
    'MCOPSTN1' = 'Microsoft 365 International Calling Plan'
    'EMS' = 'Enterprise Mobility + Security'
    'INTUNE_A' = 'Microsoft Intune'
    'RIGHTSMANAGEMENT' = 'Azure Rights Management'
    'PROJECTPROFESSIONAL' = 'Project Professional'
    'PROJECTONLINE_PLAN_1' = 'Project Online Plan 1'
    'PROJECTONLINE_PLAN_2' = 'Project Online Plan 2'
    'VISIOCLIENT' = 'Visio Pro for Office 365'
    'POWER_BI_PRO' = 'Power BI Pro'
    'POWER_BI_PREMIUM' = 'Power BI Premium'
    'FLOW_FREE' = 'Microsoft Flow Free'
    'POWERAPPS_VIRAL' = 'Power Apps'
    'STREAM' = 'Microsoft Stream'
    'TEAMS1' = 'Microsoft Teams'
    'TEAMS_EXPLORATORY' = 'Microsoft Teams Exploratory'
    'ATP_ENTERPRISE' = 'Microsoft Defender for Office 365'
    'AAD_PREMIUM' = 'Azure Active Directory Premium'
    'AAD_PREMIUM_P2' = 'Azure Active Directory Premium P2'
    'MFA_PREMIUM' = 'Azure Multi-Factor Authentication'
    'ADALLOM_STANDALONE' = 'Microsoft Cloud App Security'
    'WINDOWS_STORE' = 'Windows Store for Business'
    'WINDOWS_STORE_EDU' = 'Windows Store for Education'
    'WINDOWS_STORE_EDU_FACULTY' = 'Windows Store for Education Faculty'
    'WINDOWS_STORE_EDU_STUDENT' = 'Windows Store for Education Student'
    'WINDOWS_STORE_EDU_STUDENT_USE_BENEFIT' = 'Windows Store for Education Student Use Benefit'
    'WINDOWS_STORE_EDU_FACULTY_USE_BENEFIT' = 'Windows Store for Education Faculty Use Benefit'
    'WINDOWS_STORE_EDU_USE_BENEFIT' = 'Windows Store for Education Use Benefit'
    'WINDOWS_STORE_EDU_STUDENT_USE_BENEFIT_FACULTY' = 'Windows Store for Education Student Use Benefit Faculty'
    'WINDOWS_STORE_EDU_STUDENT_USE_BENEFIT_STUDENT' = 'Windows Store for Education Student Use Benefit Student'
    'WINDOWS_STORE_EDU_FACULTY_USE_BENEFIT_FACULTY' = 'Windows Store for Education Faculty Use Benefit Faculty'
    'WINDOWS_STORE_EDU_FACULTY_USE_BENEFIT_STUDENT' = 'Windows Store for Education Faculty Use Benefit Student'
    'WINDOWS_STORE_EDU_USE_BENEFIT_FACULTY' = 'Windows Store for Education Use Benefit Faculty'
    'WINDOWS_STORE_EDU_USE_BENEFIT_STUDENT' = 'Windows Store for Education Use Benefit Student'
}

# Function to get friendly license name
function Get-FriendlyLicenseName {
    param (
        [string]$SkuPartNumber
    )
    if ($licenseMapping.ContainsKey($SkuPartNumber)) {
        return $licenseMapping[$SkuPartNumber]
    }
    return $SkuPartNumber
}

# Function to get all users with their licenses
function Get-UserLicenses {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser[]]$Users
    )

    $userLicenses = @()
    $totalUsers = $Users.Count
    $currentUser = 0

    foreach ($user in $Users) {
        $currentUser++
        $progress = ($currentUser / $totalUsers) * 100
        Write-Progress -Activity "Processing Users" -Status "Processing $($user.UserPrincipalName)" -PercentComplete $progress

        $userLicense = [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            DisplayName = $user.DisplayName
            AccountEnabled = $user.AccountEnabled
            Licenses = @()
            AssignmentMethod = "Direct"
        }

        try {
            # Get user's licenses
            $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id
            $userLicense.Licenses = $licenseDetails | 
                Select-Object SkuId, SkuPartNumber, @{Name='ServicePlans'; Expression={$_.ServicePlans.ServicePlanName}}

            # Check if licenses are assigned via groups
            $groupAssignments = Get-MgUserMemberOf -UserId $user.Id | 
                Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' }
            
            if ($groupAssignments) {
                $userLicense.AssignmentMethod = "Group"
            }
        }
        catch {
            Write-Warning "Error processing user $($user.UserPrincipalName): $_"
            $userLicense.Licenses = @()
        }

        $userLicenses += $userLicense
    }
    Write-Progress -Activity "Processing Users" -Completed
    return $userLicenses
}

# Function to generate HTML report
function New-LicenseReport {
    param (
        [Parameter(Mandatory = $true)]
        [array]$UserLicenses
    )

    Write-Progress -Activity "Generating Report" -Status "Creating HTML Report"

    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Microsoft 365 License Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1, h2 {
            color: #0078d4;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        th, td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: #f2f2f2;
            font-weight: bold;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .warning {
            background-color: #fff3cd;
            color: #856404;
            padding: 10px;
            border-radius: 4px;
            margin: 10px 0;
        }
        .disabled-user {
            background-color: #f8d7da;
        }
        .license-list {
            list-style-type: none;
            padding: 0;
            margin: 0;
        }
        .license-item {
            margin-bottom: 5px;
            padding: 5px;
            background-color: #e9ecef;
            border-radius: 3px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Microsoft 365 License Report</h1>
        <p>Generated on: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>

        <h2>All Users and Their Licenses</h2>
        <table>
            <tr>
                <th>User</th>
                <th>Display Name</th>
                <th>Licenses</th>
                <th>Assignment Method</th>
                <th>Status</th>
            </tr>
"@

    foreach ($user in $UserLicenses) {
        $rowClass = if (-not $user.AccountEnabled) { "class='disabled-user'" } else { "" }
        $friendlyLicenses = $user.Licenses.SkuPartNumber | ForEach-Object { Get-FriendlyLicenseName $_ }
        $status = if (-not $user.AccountEnabled) { "Disabled" } else { "Active" }
        $html += @"
            <tr $rowClass>
                <td>$($user.UserPrincipalName)</td>
                <td>$($user.DisplayName)</td>
                <td><ul class="license-list">$($friendlyLicenses | ForEach-Object { "<li class='license-item'>$_</li>" })</ul></td>
                <td>$($user.AssignmentMethod)</td>
                <td>$status</td>
            </tr>
"@
    }

    $disabledUsers = $UserLicenses | Where-Object { -not $_.AccountEnabled -and $_.Licenses.Count -gt 0 }
    
    $html += @"
        </table>

        <h2>Disabled Users with Licenses</h2>
        <div class="warning">
            <p>The following users are disabled but still have licenses assigned:</p>
            <ul>
"@

    foreach ($user in $disabledUsers) {
        $friendlyLicenses = $user.Licenses.SkuPartNumber | ForEach-Object { Get-FriendlyLicenseName $_ }
        $html += @"
                <li>$($user.UserPrincipalName) - $($friendlyLicenses -join ', ')</li>
"@
    }

    $html += @"
            </ul>
        </div>

        <h2>License Recommendations</h2>
        <div class="warning">
            <p>Consider the following recommendations:</p>
            <ul>
                <li>Review and remove licenses from disabled users</li>
                <li>Consider using group-based licensing for easier management</li>
                <li>Regularly audit license assignments to ensure they match user needs</li>
            </ul>
        </div>
    </div>
</body>
</html>
"@

    Write-Progress -Activity "Generating Report" -Completed
    return $html
}

# Main script execution
try {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    # Connect to Microsoft Graph with minimal required scopes
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome

    Write-Host "Fetching user information..." -ForegroundColor Cyan
    # Get all users with minimal required properties
    $users = Get-MgUser -All -Property "id,userPrincipalName,displayName,accountEnabled"

    Write-Host "Processing license information..." -ForegroundColor Cyan
    # Get license information
    $userLicenses = Get-UserLicenses -Users $users

    Write-Host "Generating report..." -ForegroundColor Cyan
    # Generate and save HTML report
    $htmlReport = New-LicenseReport -UserLicenses $userLicenses
    $reportPath = "M365LicenseReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $htmlReport | Out-File -FilePath $reportPath -Encoding UTF8

    Write-Host "Report generated successfully: $reportPath" -ForegroundColor Green
    Write-Host "Opening report in default browser..." -ForegroundColor Cyan
    Start-Process $reportPath
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    # Disconnect from Microsoft Graph
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph
} 