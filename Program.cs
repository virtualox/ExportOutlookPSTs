using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Win32;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ExportOutlookPSTs
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check if sharePath is provided as an argument
            if (args.Length == 0)
            {
                // Log error message to a default location
                string defaultErrorLog = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ExportOutlookPSTs_ErrorLog.txt");
                string errorMessage = $"[{DateTime.Now}] - {Environment.UserName} - Error: No sharePath provided as a command-line argument.";
                LogError(defaultErrorLog, errorMessage);
                return;
            }

            string sharePath = args[0];

            // Validate sharePath format
            if (!sharePath.StartsWith(@"\\"))
            {
                string defaultErrorLog = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ExportOutlookPSTs_ErrorLog.txt");
                string errorMessage = $"[{DateTime.Now}] - {Environment.UserName} - Error: Invalid network share path format: '{sharePath}'. Ensure it starts with '\\\\'.";
                LogError(defaultErrorLog, errorMessage);
                return;
            }

            string username = Environment.UserName;
            string logFile = Path.Combine(sharePath, $"{username}.txt");
            string errorLog = Path.Combine(sharePath, "ErrorLog.txt");
            string noPstLog = Path.Combine(sharePath, "_NoPST.txt");

            try
            {
                // Check if the network share is accessible
                if (!Directory.Exists(sharePath))
                {
                    throw new DirectoryNotFoundException($"The network share path '{sharePath}' is not accessible.");
                }

                string logEntry;

                // Check if an Outlook profile exists
                if (!OutlookProfileExists())
                {
                    // No Outlook profiles found
                    logEntry = $"{username} - No Outlook profile - checked: {DateTime.Now}";
                    UpdateNoPstLog(noPstLog, username, logEntry);
                    return;
                }

                // Initialize Outlook COM object
                Outlook.Application outlookApp = null;
                Outlook.NameSpace mapiNamespace = null;

                try
                {
                    outlookApp = new Outlook.Application();

                    // Get MAPI namespace without displaying UI
                    mapiNamespace = outlookApp.GetNamespace("MAPI");
                    mapiNamespace.Logon("", "", Missing.Value, false);
                }
                catch (COMException comEx)
                {
                    throw new Exception($"COM Exception: Failed to create Outlook COM object or get MAPI namespace: {comEx.Message}", comEx);
                }
                catch (Exception ex)
                {
                    throw new Exception($"General Exception: Failed to create Outlook COM object or get MAPI namespace: {ex.Message}", ex);
                }

                // Collect PST paths
                var pstPaths = mapiNamespace.Stores
                    .Cast<Outlook.Store>()
                    .Select(store => store.FilePath)
                    .Where(path => !string.IsNullOrEmpty(path) && File.Exists(path) && path.EndsWith(".pst", StringComparison.OrdinalIgnoreCase))
                    .Select(NormalizePath)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                // Log off MAPI namespace
                mapiNamespace.Logoff();

                // Release COM objects
                Marshal.ReleaseComObject(mapiNamespace);
                mapiNamespace = null;
                Marshal.ReleaseComObject(outlookApp);
                outlookApp = null;

                if (pstPaths.Count == 0)
                {
                    // No PST files found
                    logEntry = $"{username} - No PST files - checked: {DateTime.Now}";
                    UpdateNoPstLog(noPstLog, username, logEntry);
                    return;
                }

                // Read existing log entries and normalize them
                var existingEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (File.Exists(logFile))
                {
                    foreach (var line in File.ReadAllLines(logFile))
                    {
                        var normalized = NormalizePath(line.Trim());
                        if (!string.IsNullOrEmpty(normalized))
                            existingEntries.Add(normalized);
                    }
                }

                // Determine new entries not yet in the log
                var newEntries = pstPaths.Where(p => !existingEntries.Contains(p)).ToList();

                if (newEntries.Count > 0)
                {
                    // Add new entries to the log
                    File.AppendAllLines(logFile, newEntries);
                }
                // If there are no new entries, do nothing

                // Remove user from _NoPST.txt if they now have PST files
                RemoveUserFromNoPstLog(noPstLog, username);
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                string errorMessage = $"[{DateTime.Now}] - {username} - Unexpected error in main script: {ex.Message}";
                LogError(errorLog, errorMessage);
            }
        }

        /// <summary>
        /// Checks if an Outlook profile exists by inspecting the registry.
        /// </summary>
        /// <returns>True if a profile exists, otherwise false.</returns>
        private static bool OutlookProfileExists()
        {
            // Possible Outlook versions
            string[] outlookVersions = { "16.0", "15.0", "14.0", "12.0", "11.0" };

            foreach (var version in outlookVersions)
            {
                string profileKeyPath = $@"Software\Microsoft\Office\{version}\Outlook\Profiles";
                using (RegistryKey profileKey = Registry.CurrentUser.OpenSubKey(profileKeyPath))
                {
                    if (profileKey != null)
                    {
                        string[] profiles = profileKey.GetSubKeyNames();
                        if (profiles != null && profiles.Length > 0)
                        {
                            // Profiles exist
                            return true;
                        }
                    }
                }
            }
            // No profiles found
            return false;
        }

        /// <summary>
        /// Normalizes the file path by removing the '\\?\UNC\' prefix if present.
        /// </summary>
        /// <param name="path">The original file path.</param>
        /// <returns>The normalized file path.</returns>
        private static string NormalizePath(string path)
        {
            if (path.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
            {
                return @"\\" + path.Substring(8);
            }
            return path;
        }

        /// <summary>
        /// Logs users without PST files or Outlook profiles to the _NoPST.txt file without duplicates.
        /// Updates the existing entry if the user is already in the log.
        /// </summary>
        /// <param name="noPstLogPath">Path to the NoPST log file.</param>
        /// <param name="username">The username to log.</param>
        /// <param name="logEntry">The log entry to add or update.</param>
        private static void UpdateNoPstLog(string noPstLogPath, string username, string logEntry)
        {
            try
            {
                // Ensure the directory exists
                string directory = Path.GetDirectoryName(noPstLogPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var logLines = new List<string>();

                // Read existing log entries
                if (File.Exists(noPstLogPath))
                {
                    logLines = File.ReadAllLines(noPstLogPath).ToList();

                    // Check if user already exists in the log
                    bool userFound = false;
                    for (int i = 0; i < logLines.Count; i++)
                    {
                        if (logLines[i].StartsWith(username + " -", StringComparison.OrdinalIgnoreCase))
                        {
                            // Update the existing entry
                            logLines[i] = logEntry;
                            userFound = true;
                            break;
                        }
                    }

                    if (!userFound)
                    {
                        // Add new entry
                        logLines.Add(logEntry);
                    }
                }
                else
                {
                    // Log file does not exist, create new entry
                    logLines.Add(logEntry);
                }

                // Write updated log entries back to the file
                File.WriteAllLines(noPstLogPath, logLines);
            }
            catch
            {
                // If logging fails, do nothing since the application runs hidden
            }
        }

        /// <summary>
        /// Removes the user from the _NoPST.txt log file if they now have PST files.
        /// </summary>
        /// <param name="noPstLogPath">Path to the NoPST log file.</param>
        /// <param name="username">The username to remove.</param>
        private static void RemoveUserFromNoPstLog(string noPstLogPath, string username)
        {
            try
            {
                if (!File.Exists(noPstLogPath))
                {
                    return;
                }

                var logLines = File.ReadAllLines(noPstLogPath).ToList();
                bool userFound = false;

                // Remove the user's entry if it exists
                for (int i = 0; i < logLines.Count; i++)
                {
                    if (logLines[i].StartsWith(username + " -", StringComparison.OrdinalIgnoreCase))
                    {
                        logLines.RemoveAt(i);
                        userFound = true;
                        break;
                    }
                }

                if (userFound)
                {
                    // Write updated log entries back to the file
                    File.WriteAllLines(noPstLogPath, logLines);
                }
            }
            catch
            {
                // If operation fails, do nothing since the application runs hidden
            }
        }

        /// <summary>
        /// Logs error messages to the specified error log file.
        /// </summary>
        /// <param name="errorLogPath">Path to the error log file.</param>
        /// <param name="message">Error message to log.</param>
        private static void LogError(string errorLogPath, string message)
        {
            try
            {
                // Ensure the directory exists
                string directory = Path.GetDirectoryName(errorLogPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Append the error message to the error log file
                File.AppendAllText(errorLogPath, message + Environment.NewLine);
            }
            catch
            {
                // If logging fails, do nothing since the application runs hidden
            }
        }
    }
}