# Microsoft 365 License Report Generator

A PowerShell script that generates a detailed HTML report of Microsoft 365 user licenses, including disabled users with licenses and optimization recommendations.

## Features

- Fetches user license information via Microsoft Graph API
- Generates a clean, well-designed HTML report
- Identifies disabled users with active licenses
- Provides license management recommendations
- Shows friendly license names instead of SKU IDs
- Displays license assignment methods (Direct/Group-based)
- Includes progress indicators for better user experience

## Prerequisites

- PowerShell 7.0 or later
- Microsoft Graph PowerShell modules:
  - Microsoft.Graph.Authentication
  - Microsoft.Graph.Users
  - Microsoft.Graph.Users.Actions

## Installation

1. Install the required Microsoft Graph PowerShell modules:
```powershell
Install-Module Microsoft.Graph.Authentication
Install-Module Microsoft.Graph.Users
Install-Module Microsoft.Graph.Users.Actions
```

2. Clone this repository:
```powershell
git clone https://github.com/yourusername/Licences-MS365.git
cd Licences-MS365
```

## Usage

1. Run the script:
```powershell
.\Get-M365Licenses.ps1
```

2. When prompted, sign in with your Microsoft 365 administrator account

3. The script will:
   - Connect to Microsoft Graph
   - Fetch user and license information
   - Generate an HTML report
   - Open the report in your default browser

## Required Permissions

The script requires the following Microsoft Graph API permissions:
- User.Read.All
- Directory.Read.All

## Output

The script generates an HTML report (`M365LicenseReport_YYYYMMDD_HHMMSS.html`) containing:
- List of all users and their assigned licenses
- Disabled users with active licenses
- License assignment methods
- Optimization recommendations

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 