# AccountTester [![Security Checks](https://github.com/Miiraak/Account-Tester/actions/workflows/security-checks.yml/badge.svg)](https://github.com/Miiraak/Account-Tester/actions/workflows/security-checks.yml)
<p align="center">
    <img src=".github/App.gif">
  
## Description
AccountTester is a Windows Forms application (C#) designed to test various aspects of user accounts on a system. It performs checks such as Internet connectivity, access rights to network drives, Office presence and permissions, as well as printer availability. A detailed test report export is now available.

## Features
- **Internet Connectivity Test:** Verifies if the computer can access the Internet by sending a request to Google.
- **Network Drive Access Test:** Attempts to create and delete a test file on each mapped network drive to check write permissions.
- **Office Version Detection:** Searches for the presence of Microsoft Office via the system registry.
- **Office Read/Write Permissions Test:** Creates, edits, and reads a Word document to verify user permissions with interop.word.
- **Installed Printers List:** Displays all printers available on the system.
- **Detailed Test Report:** Export the full test report in `.txt`, `.json`, `.log`, `.csv` or `.xml`. Use `.zip` for all in one.
- **Language:** Now you can change the language of the application and the report ! 
    - Langage available :
        - `EN` - `100%`
        - `FR` - `100%`

### Features in development
| Nom | Desc. |
|---|---|
| **Report export formats** | Add support for `.csv`, `.json` and `.xml` export formats. | 
| **Improved Interface** | 	Enhancing the UI for better readability and user experience. |
| **...** | ... |

## Prerequisites
Before running the project, make sure you have the following:

- Windows with .NET Framework installed.
- Microsoft Office installed (for Word-related tests).
- Sufficient access rights to test network drives and the Windows registry.
> The software is designed to work with at least the rights of a local non-admin user account.

## Usage
1. Launch the application.
2. Click the Start button.
3. Wait for the tests to complete.
4. View the results in the log area.
5. Export results if desired.
6. Choose a name (a default name will be used if left empty).
7. Choose the save location.
8. Select the export format.

## Contributing
Contributions are welcome! To contribute to this project, please follow these steps:

1. Fork the repository.
2. Create a new branch for your feature (git checkout -b my-new-feature).
3. Make your changes.
4. Commit your changes (git commit -m 'Add my new feature').
5. Push your branch (git push origin my-new-feature).
6. Open a Pull Request.

## Issues and Suggestions
If you encounter issues or have suggestions to improve the project, please use the [GitHub issue tracker](https://github.com/Miiraak/Account-Tester/issues).

## License
This project is not licensed. All rights reserved.

## Authors
- [**Miiraak**](https://github.com/miiraak) - *Lead Developer*

---
