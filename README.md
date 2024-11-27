# ExportOutlookPSTs

A utility to export and log Outlook PST file paths to a central network share.

## Features

- Retrieves PST file paths from Outlook profiles.
- Normalizes paths to avoid duplicates.
- Logs PST paths to per-user log files on a network share.
- Logs users without PST files in a separate log.
- Handles errors gracefully and logs them.
- Runs hidden without user intervention.

## Requirements

- .NET Framework 4.7.2 or later.
- Microsoft Outlook installed and configured.
- Access to a network share for logging.

## Installation

1. **Clone the repository:**

   ```bash
   git clone https://github.com/virtualox/ExportOutlookPSTs.git
   ```

2. **Open the solution in Visual Studio.**

3. **Build the project in Release mode.**

## Usage

Run the executable with the network share path as a command-line argument:
   
   ```bash
   ExportOutlookPSTs.exe "\\YourServer\YourShare$"
   ```

## Logging Behavior

- User Log Files (`<username>.txt`):
  - Contains the list of PST file paths associated with the user.
  - Avoids duplicate entries.
- No PST Log (`_NoPST.txt`):
  - Records users who do not have any PST files.
  - Entries are in the format: `<username> - checked: <date and time>`
- Error Log (`ErrorLog.txt`):
  - Captures any errors encountered during execution.
