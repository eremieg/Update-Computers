# Windows Computers Update and Report Automation

This PowerShell script automates the process of updating and generating a report for multiple Windows computers. It reads a CSV file containing computer names, retrieves their MAC addresses using Active Directory, wakes up the computers using Wake-on-LAN, and installs updates using the PSWindowsUpdate module. Additionally, it retrieves the last logged-on user for each computer. Finally, it generates an HTML report summarizing the update status and logs its activities to a separate log file for detailed review.

## Prerequisites

1. PowerShell 5.1 or later
2. Active Directory PowerShell module
3. PSWindowsUpdate module

## Usage

1. Prepare a CSV file containing computer names, with a header named 'ComputerName'.
2. Run the script in PowerShell.
3. When prompted, provide the path to the CSV file.
4. The script will wake up the computers, install updates, and generate an HTML report with the update status.

## Notes

- The script uses Wake-on-LAN to wake up the computers; ensure that the target computers have Wake-on-LAN enabled.
- The script must be run with administrative privileges to access Active Directory and perform updates on remote computers.
- The script uses PowerShell jobs to process multiple computers in parallel for improved performance.

## Example CSV File

Create a CSV file containing a list of computer names, with a header named 'ComputerName'. For example:

ComputerName  
PC1  
PC2  
PC3  

## Generated HTML Report

The generated HTML report contains the following sections:

1. Successfully Updated Computers: Lists computers that were successfully updated, the installed updates, and the last logged-on user.
2. Computers That Failed to Wake Up: Lists computers that failed to wake up after multiple attempts.
3. Computers with Update Errors: Lists computers that encountered errors during the update process.

## Log File

The log file, named `UpdateComputers.log`, is created in the same directory as the script. It contains detailed information about the script's activities, including attempts to wake up computers, update progress, and any errors encountered. Review the log file for additional information in case of errors or issues with the update process.

## License

This script is free to use, modify, and distribute under the terms of the [MIT License](https://opensource.org/licenses/MIT).
