using Microsoft.Win32;
using mkWritable.Properties;
using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Linq;

namespace mkWritable
{
    class Program
    {
        private static string argument;

        [STAThread]
        static void Main(string[] args)
        {
            // Check if no argument is provided, show an "About" dialog box
            if (args.Length > 0)
            {
                argument = args[0].ToLower();
                switch (argument)
                {
                    case "/register":
                        RegisterContextEntry();
                        break;
                    case "/unregister":
                        UnregisterContextEntry();
                        break;
                    default:
                        Application.EnableVisualStyles();
                        Application.SetCompatibleTextRenderingDefault(false);
                        ProcessFile(argument);
                        break;
                }
            }
            else
            {
                ShowAboutBox();
                return;
            }

        }

        static void ProcessFile(string filePath)
        {
            // Check if the file exists
            if (!File.Exists(filePath))
            {
                ShowBalloonNotification("mkWritable", String.Format(Resources.TheFile0DoesNotExist, filePath));
                return;
            }

            // Determine if it's an Excel or Word file
            string fileExtension = Path.GetExtension(filePath).ToLower();
            if (fileExtension == ".xlsx")
            {
                ProcessExcelFile(filePath);
            }
            else if (fileExtension == ".docx")
            {
                ProcessWordFile(filePath);
            }
            else
            {
                ShowBalloonNotification("mkWritable", Resources.UnsupportedFileTypeOnlyXlsxAndDocxAreSupported);
            }
        }
        static void ProcessExcelFile(string excelFilePath)
        {
            string extractedPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(excelFilePath) + "_extracted");

            try
            {
                // 1. Extract the .zip file
                ZipFile.ExtractToDirectory(excelFilePath, extractedPath);

                // 2. Iterate through all sheet*.xml files in xl/worksheets/
                string worksheetsFolder = Path.Combine(extractedPath, "xl", "worksheets");
                if (!Directory.Exists(worksheetsFolder))
                {
                    ShowBalloonNotification("mkWritable", Resources.TheWorksheetsFolderCouldNotBeFound);
                    return;
                }

                var sheetFiles = Directory.GetFiles(worksheetsFolder, "sheet*.xml");

                if (sheetFiles.Length == 0)
                {
                    ShowBalloonNotification("mkWritable", Resources.NoSheetXmlFilesFound);
                    return;
                }

                foreach (var sheetFile in sheetFiles)
                {
                    XDocument sheetXml = XDocument.Load(sheetFile);

                    // Dynamically read the default namespace from the root element
                    XNamespace ns = sheetXml.Root.GetDefaultNamespace();

                    // Remove the <sheetProtection> element
                    var sheetProtectionElement = sheetXml.Descendants(ns + "sheetProtection").FirstOrDefault();

                    if (sheetProtectionElement != null)
                    {
                        sheetProtectionElement.Remove();
                        ShowBalloonNotification("mkWritable", String.Format(Resources.WriteProtectionFoundAndRemovedIn0, excelFilePath));
                    }
                    else
                    {
                        ShowBalloonNotification("mkWritable", String.Format(Resources.NoWriteProtectionFoundIn0, excelFilePath));
                    }

                    // Save the modified sheet XML
                    sheetXml.Save(sheetFile);
                }

                // 5. Repackage the files into a new .zip file and rename it back to .xlsx
                string newExcelFilePath = Path.ChangeExtension(excelFilePath, ".modified.xlsx");
                ZipFile.CreateFromDirectory(extractedPath, newExcelFilePath);

                ShowBalloonNotification("mkWritable", Resources.ProcessCompletedNewExcelFileSavedAt + newExcelFilePath);
            }
            finally
            {
                Cleanup(extractedPath);
            }
        }

        static void ProcessWordFile(string docxFilePath)
        {
            string extractedPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(docxFilePath) + "_extracted");

            try
            {
                // 1. Extract the .zip file
                ZipFile.ExtractToDirectory(docxFilePath, extractedPath);

                // 2. Locate the settings.xml file inside the word folder
                string settingsFilePath = Path.Combine(extractedPath, "word", "settings.xml");
                if (!File.Exists(settingsFilePath))
                {
                    ShowBalloonNotification("mkWritable", Resources.TheSettingsXmlFileCouldNotBeFound);
                    return;
                }

                // 3. Directly manipulate the XML content as a string
                string xmlContent = File.ReadAllText(settingsFilePath);
                if (xmlContent.Contains("<w:documentProtection"))
                {
                    // Remove all occurrences of <w:documentProtection ... /> using regex
                    string pattern = @"<w:documentProtection[^>]*?\/>";
                    xmlContent = System.Text.RegularExpressions.Regex.Replace(xmlContent, pattern, string.Empty);

                    // Save the modified content back to settings.xml
                    File.WriteAllText(settingsFilePath, xmlContent);
                    ShowBalloonNotification("mkWritable", String.Format(Resources.WriteProtectionRemovedFrom0, Path.GetFileName(docxFilePath)));
                }
                else
                {
                    ShowBalloonNotification("mkWritable", String.Format(Resources.NoWriteProtectionFoundIn0, Path.GetFileName(docxFilePath)));
                }


                // 4. Repackage the modified document
                string newDocxFilePath = Path.ChangeExtension(docxFilePath, ".modified.docx");

                // Create a new .zip archive from the extracted directory
                ZipFile.CreateFromDirectory(extractedPath, newDocxFilePath);

                ShowBalloonNotification("mkWritable", Resources.ProcessCompletedNewWordFileSavedAt + newDocxFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(Resources.AnErrorOccurred + ex.Message);
            }
            finally
            {
                // Cleanup temporary files
                Cleanup(extractedPath);
            }
        }

        // Method to clean up temporary files
        static void Cleanup(string extractedPath)
        {
            if (Directory.Exists(extractedPath))
            {
                Directory.Delete(extractedPath, true);
            }
        }

        // Method to show an "About" dialog box if no parameter is provided
        static void ShowAboutBox()
        {
            string progName = Assembly.GetExecutingAssembly().GetName().Name;
            Version shortVersion = Assembly.GetExecutingAssembly().GetName().Version;
            string about = string.Format(Resources.About, progName, $" {shortVersion.Major}.{shortVersion.Minor}.{shortVersion.Build}");

            // Create a hidden form to show the MessageBox without showing a taskbar icon
            using (Form hiddenForm = new Form()
            {
                ShowInTaskbar = true, // Ensure the form is in the taskbar to show the icon althought the form is hidden
                Icon = Resources.mkWritable,
                Opacity = 0, // Set the form's opacity to 0, making it completely invisible
                Width = 0,
                Height = 0,
                FormBorderStyle = FormBorderStyle.None, // Set the border style to None
                StartPosition = FormStartPosition.CenterScreen // Optional: Center the message box on the screen
            })
            {
                hiddenForm.Load += (s, e) => hiddenForm.Hide(); // Hide the form as soon as it loads

                // Display the MessageBox
                MessageBox.Show(hiddenForm, about, Resources.AboutMkWritable, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // Method to show balloon notification in GUI mode
        static void ShowBalloonNotification(string title, string message)
        {
            using (NotifyIcon notifyIcon = new NotifyIcon())
            {
                notifyIcon.Visible = true;
                notifyIcon.Icon = SystemIcons.Information; // Using a system icon
                notifyIcon.BalloonTipTitle = title;
                notifyIcon.BalloonTipText = message;
                notifyIcon.ShowBalloonTip(3000); // Show for 3 seconds

                // Sleep for a short time to ensure the notification is visible
                System.Threading.Thread.Sleep(4000);
            }
        }

        // Create registry entries for both Excel.Sheet.12 and Word.Document.12 context menus
        static void RegisterContextEntry()
        {
            string[] progIds = { "Excel.Sheet.12", "Word.Document.12" };

            foreach (string progId in progIds)
            {
                // Check if the shell key exists, create it if it doesn't
                using (RegistryKey shellKey = Registry.CurrentUser.OpenSubKey($"Software\\Classes\\{progId}\\shell", writable: true) ??
                                              Registry.CurrentUser.CreateSubKey($"Software\\Classes\\{progId}\\shell"))
                {
                    // Create mkWritable entry under the shell key
                    using (RegistryKey key = shellKey.CreateSubKey("mkWritable"))
                    {
                        key.SetValue("", Resources.MakeWritable);
                        key.SetValue("NoWorkingDirectory", "");
                        key.SetValue("Position", "bottom");
                        key.SetValue("Icon", Application.ExecutablePath + ",0");
                    }

                    // Create the command subkey and set the command to execute the app
                    using (RegistryKey commandKey = shellKey.CreateSubKey(@"mkWritable\command"))
                    {
                        commandKey.SetValue("", $"\"{Application.ExecutablePath}\" \"%1\"");
                    }
                }
            }
            ShowBalloonNotification("mkWritable", Resources.ApplicationRegisteredInUsersContextMenu);
        }

        // Unregister the context menu from both Excel.Sheet.12 and Word.Document.12
        static void UnregisterContextEntry()
        {
            string[] progIds = { "Excel.Sheet.12", "Word.Document.12" };

            foreach (string progId in progIds)
            {
                // Delete mkWritable key
                Registry.CurrentUser.DeleteSubKeyTree($"Software\\Classes\\{progId}\\shell\\mkWritable", false);

                // Check if the shell key is now empty, and delete it if so
                using (RegistryKey shellKey = Registry.CurrentUser.OpenSubKey($"Software\\Classes\\{progId}\\shell", writable: true))
                {
                    if (shellKey != null && shellKey.SubKeyCount == 0)
                    {
                        Registry.CurrentUser.DeleteSubKeyTree($"Software\\Classes\\{progId}\\shell", false);
                    }
                }
            }

            ShowBalloonNotification("mkWritable", Resources.ApplicationUnregisteredFromUsersContextMenu);
        }
    }
}
