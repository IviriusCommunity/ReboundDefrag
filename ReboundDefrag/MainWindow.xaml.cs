using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.Win32.SafeHandles;
using ReboundDefrag.Helpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Windows.ApplicationModel;
using Windows.ApplicationModel.Core;
using WinUIEx;
using WinUIEx.Messaging;

#nullable enable

namespace ReboundDefrag
{
    public sealed partial class MainWindow : WindowEx
    {
        public MainWindow()
        {
            this?.InitializeComponent();
        }

        public async Task LoadAppAsync()
        {
            // Set window backdrop to Mica for a modern translucent effect
            SystemBackdrop = new MicaBackdrop();

            // Set the window title
            Title = "Optimize Drives - ALPHA v0.0.3";

            // Window customization
            IsMaximizable = false;
            this.SetWindowSize(800, 670);
            IsResizable = false;
            AppWindow.DefaultTitleBarShouldMatchAppModeTheme = true;
            this.CenterOnScreen();
            this.SetIcon(@$"{AppContext.BaseDirectory}/Assets/ReboundDefrag.ico");

            // Load data based on the current state of 'AdvancedView'
            await LoadData(AdvancedView.IsOn);

            // Begin monitoring window messages (such as device changes)
            var hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WindowMessageMonitor mon = new(hWnd);

            // Subscribe to the WindowMessageReceived event
            mon.WindowMessageReceived += MessageReceived;

            // Set up a timer to periodically refresh the message listener every 5 seconds
            var timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(5) // Timer interval set to 5 seconds
            };

            // Timer event handler
            timer.Tick += (sender, e) =>
            {
                mon.WindowMessageReceived += MessageReceived; // Re-subscribe to the event to keep monitoring
            };

            // Start the timer
            timer.Start();

            // Check if the application is running with administrator privileges
            IsAdministrator();
        }

        private async void MessageReceived(object? sender, WindowMessageEventArgs e)
        {
            // Handle incoming messages
            switch (e.Message.MessageId)
            {
                default:
                    {
                        // No relevant message, break
                        break;
                    }
                case Win32Helper.WM_DEVICECHANGE:
                    {
                        // Handle specific device changes
                        switch ((int)e.Message.WParam)
                        {
                            case Win32Helper.DBT_DEVICEARRIVAL:
                                {
                                    // Device or partition inserted
                                    MyListView.ItemsSource = null; // Clear existing list
                                    await LoadData(AdvancedView.IsOn); // Reload data based on AdvancedView state
                                    break;
                                }
                            case Win32Helper.DBT_DEVICEREMOVECOMPLETE:
                                {
                                    // Device or partition removed
                                    MyListView.ItemsSource = null; // Clear existing list
                                    await LoadData(AdvancedView.IsOn); // Reload data based on AdvancedView state
                                    break;
                                }
                            default:
                                {
                                    // Handle any other device action
                                    break;
                                }
                        }
                        break;
                    }
            }
        }

        public class DiskItem : Item
        {
            public string? DriveLetter { get; set; }
            public string? MediaType { get; set; }
            public int ProgressValue { get; set; }
        }

        public bool IsAdministrator()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            if (principal.IsInRole(WindowsBuiltInRole.Administrator))
            {
                Admin1.Visibility = Visibility.Collapsed;
                Admin2.Visibility = Visibility.Collapsed;
                Admin3.Visibility = Visibility.Collapsed;
                Admin4.Visibility = Visibility.Collapsed;
            }
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private void ListviewSelectionChange(object sender, SelectionChangedEventArgs args)
        {
            LoadSelectedItemInfo(GetStatus());
        }

        public string GetLastOptimizeDate()
        {
            if (MyListView.SelectedItem != null)
            {
                DiskItem? selectedItem = MyListView.SelectedItem as DiskItem;
                // Handle the selection change event

                try
                {
                    List<string> i = DefragInfo.GetEventLogEntriesForID(258);
                    return i.Last(s => s.Contains($"({selectedItem?.DriveLetter?.ToString().Remove(2, 1)})"));
                }
                catch
                {
                    return "Never....";
                }
            }
            else return "Unknown....";
        }

        public void LoadSelectedItemInfo(string status, string info = "....")
        {
            if (info == "....")
            {
                info = GetLastOptimizeDate();
            }
            else
            {
                info = "Unknown....";
            }
            if (MyListView.SelectedItem != null)
            {
                DiskItem? selectedItem = MyListView.SelectedItem as DiskItem;
                DetailsBar.Title = selectedItem?.Name;
                DetailsBar.Message = $"Media type: {selectedItem?.MediaType}\nLast analyzed or optimized: {info[..^4]}\nCurrent status: {status}";
                DetailsBar.Severity = InfoBarSeverity.Informational;
                OptimizeButton.IsEnabled = true;
                if (status.Contains("Needs optimization"))
                {
                    DetailsBar.Severity = InfoBarSeverity.Warning;
                }
                if (selectedItem?.MediaType == "CD-ROM")
                {
                    DetailsBar.Message = "Media type: CD-ROM\nLast analyzed or optimized: Never\nCurrent status: cannot be optimized";
                    DetailsBar.Severity = InfoBarSeverity.Error;
                    VisualStateManager.GoToState(OptimizeButton, "Disabled", true);
                    OptimizeButton.IsEnabled = false;
                }
                if (selectedItem?.Name == "EFI System Partition")
                {
                    DetailsBar.Message = $"Media type: {selectedItem.MediaType}\nLast analyzed or optimized: Never\nCurrent status: cannot be optimized (EFI System Partition)";
                    DetailsBar.Severity = InfoBarSeverity.Informational;
                }
                if (selectedItem?.Name == "Recovery Partition")
                {
                    DetailsBar.Message = $"Media type: {selectedItem.MediaType}\nLast analyzed or optimized: Never\nCurrent status: cannot be optimized (Recovery Partition)";
                    DetailsBar.Severity = InfoBarSeverity.Error;
                    VisualStateManager.GoToState(OptimizeButton, "Disabled", true);
                    OptimizeButton.IsEnabled = false;
                }
            }
        }

        public string GetStatus()
        {
            string status = string.Empty;

            try
            {
                if (MyListView.SelectedItem != null)
                {
                    List<string> i = DefragInfo.GetEventLogEntriesForID(258);
                    var selectedItem = MyListView.SelectedItem as DiskItem;

                    var selI = i.Last(s => s.Contains($"({selectedItem?.DriveLetter?.ToString().Remove(2, 1)})"));

                    var localDate = DateTime.Parse(selI[..^4]);

                    // Get the current local date and time
                    DateTime currentDate = DateTime.Now;

                    // Calculate the days passed
                    TimeSpan timeSpan = currentDate - localDate;
                    int daysPassed = timeSpan.Days;

                    if (daysPassed == 0)
                    {
                        //return $"OK (Last optimized: today)";
                        return $"OK";
                    }

                    if (daysPassed == 1)
                    {
                        //return $"OK (Last optimized: yesterday)";
                        return $"OK";
                    }

                    if (daysPassed < 50)
                    {
                        //return $"OK (Last optimized: {daysPassed} days ago)";
                        return $"OK";
                    }

                    if (daysPassed >= 50)
                    {
                        //return $"Needs optimization (Last optimized: {daysPassed} days ago)";
                        return $"Needs optimization";
                    }

                    else return "Unknown";
                }
                else
                {
                    return "Please select an item to proceed.";
                }
            }
            catch
            {
                return "Needs optimization";
            }
        }

        public void Lock(bool areItemsEnabled, string message = "")
        {
            AnalyzeButton.IsEnabled = areItemsEnabled;
            OptimizeButton.IsEnabled = areItemsEnabled;
            MyListView.IsEnabled = areItemsEnabled;
            AdvancedView.IsEnabled = areItemsEnabled;
            CurrentDisk.Visibility = areItemsEnabled == true ? Visibility.Collapsed : Visibility.Visible;
            CurrentDisk.Text = message;
            CurrentProgress.IsIndeterminate = !areItemsEnabled;
        }

        public async Task LoadData(bool loadSystemPartitions)
        {
            Lock(false, "Loading drive information...");

            // Initial delay
            // Essential for ensuring the UI loads before starting tasks
            await Task.Delay(100);

            List<DiskItem> items = [];

            // Get the logical drives bitmask
            uint drivesBitMask = Win32Helper.GetLogicalDrives();
            if (drivesBitMask == 0)
            {
                return;
            }

            for (char driveLetter = 'A'; driveLetter <= 'Z'; driveLetter++)
            {
                uint mask = 1u << (driveLetter - 'A');
                if ((drivesBitMask & mask) != 0)
                {
                    string drive = $"{driveLetter}:\\";

                    StringBuilder volumeName = new(261);
                    StringBuilder fileSystemName = new(261);
                    if (Win32Helper.GetVolumeInformation(drive, volumeName, volumeName.Capacity, out _, out _, out _, fileSystemName, fileSystemName.Capacity))
                    {
                        var newDriveLetter = drive.ToString().Remove(2, 1);
                        string mediaType = GetDriveTypeDescriptionAsync(drive);

                        if (volumeName.ToString() != string.Empty)
                        {
                            var item = new DiskItem
                            {
                                Name = $"{volumeName} ({newDriveLetter})",
                                ImagePath = "ms-appx:///Assets/Drive.png",
                                MediaType = mediaType,
                                DriveLetter = drive,
                            };
                            if (item.MediaType == "Removable")
                            {
                                item.ImagePath = "ms-appx:///Assets/DriveRemovable.png";
                            }
                            if (item.MediaType == "Unknown")
                            {
                                item.ImagePath = "ms-appx:///Assets/DriveUnknown.png";
                            }
                            if (item.MediaType == "CD-ROM")
                            {
                                item.ImagePath = "ms-appx:///Assets/DriveOptical.png";
                            }
                            if (item.DriveLetter.Contains('C'))
                            {
                                item.ImagePath = "ms-appx:///Assets/DriveWindows.png";
                            }
                            items.Add(item);
                        }
                        else
                        {
                            var item = new DiskItem
                            {
                                Name = $"({newDriveLetter})",
                                ImagePath = "ms-appx:///Assets/Drive.png",
                                MediaType = mediaType,
                                DriveLetter = drive,
                            };
                            if (item.MediaType == "Removable")
                            {
                                item.ImagePath = "ms-appx:///Assets/DriveRemovable.png";
                            }
                            if (item.MediaType == "Unknown")
                            {
                                item.ImagePath = "ms-appx:///Assets/DriveUnknown.png";
                            }
                            if (item.MediaType == "CD-ROM")
                            {
                                item.ImagePath = "ms-appx:///Assets/DriveOptical.png";
                            }
                            if (item.DriveLetter.Contains('C'))
                            {
                                item.ImagePath = "ms-appx:///Assets/DriveWindows.png";
                            }
                            items.Add(item);
                        }
                    }
                    else
                    {
                        Debug.WriteLine($"  Failed to get volume information for {drive}");
                    }
                }
            }

            if (loadSystemPartitions)
            {
                var syspart = SystemVolumes.GetSystemVolumes();

                // Add system partitions to the items list
                foreach (var result in syspart)
                {
                    string? driveType = string.Empty;
                    foreach (var diskitem in items)
                    {
                        if (diskitem.DriveLetter != null && diskitem.DriveLetter.Contains('C'))
                        {
                            driveType = diskitem.MediaType;
                        }
                    }
                    var item = new DiskItem
                    {
                        Name = result.FriendlyName,
                        ImagePath = "ms-appx:///Assets/DriveSystem.png",
                        MediaType = driveType,
                        DriveLetter = result.GUID,
                    };
                    items.Add(item);
                }
            }

            int selIndex = MyListView.SelectedIndex is not -1 ? MyListView.SelectedIndex : 0;

            // Set the list view's item source
            MyListView.ItemsSource = items;

            MyListView.SelectedIndex = selIndex >= items.Count ? items.Count - 1 : selIndex;

            Lock(true);
        }

        public static string GetDiskDriveFromLetter(string driveLetter)
        {
            try
            {
                var FULLSEARCHER = new ManagementObjectSearcher("root\\Microsoft\\Windows\\Storage", "SELECT * FROM MSFT_PhysicalDisk");
                foreach (ManagementObject queryObj in FULLSEARCHER.Get().Cast<ManagementObject>())
                {
                    string? MEDIATYPE = queryObj["MediaType"].ToString();
                    string? DEVICEID = queryObj["DeviceID"].ToString();

                    // Create a ManagementObjectSearcher to query the Win32_DiskDrive class
                    ManagementObjectSearcher searcher2 = new("SELECT * FROM Win32_DiskDrive");

                    foreach (ManagementObject disk in searcher2.Get().Cast<ManagementObject>())
                    {
                        // Get the device ID of the disk
                        string? deviceID2 = disk["DeviceID"].ToString();

                        // Create a ManagementObjectSearcher to query the Win32_LogicalDiskToPartition class
                        ManagementObjectSearcher partitionSearcher = new(
                            $"ASSOCIATORS OF {{Win32_DiskDrive.DeviceID='{deviceID2}'}} WHERE AssocClass=Win32_DiskDriveToDiskPartition");

                        foreach (ManagementObject partition in partitionSearcher.Get().Cast<ManagementObject>())
                        {
                            // Create a ManagementObjectSearcher to query the Win32_LogicalDisk class
                            ManagementObjectSearcher logicalDiskSearcher = new(
                                $"ASSOCIATORS OF {{Win32_DiskPartition.DeviceID='{partition["DeviceID"]}'}} WHERE AssocClass=Win32_LogicalDiskToPartition");

                            foreach (ManagementObject logicalDisk in logicalDiskSearcher.Get().Cast<ManagementObject>())
                            {
                                // Check if the logical disk is drive C
                                if (logicalDisk["DeviceID"].ToString() == driveLetter.Replace(@"\", ""))
                                {
                                    if (deviceID2 != null && DEVICEID != null && deviceID2.Contains(DEVICEID))
                                    {
                                        string driveType = MEDIATYPE switch
                                        {
                                            "3" => "HDD (Hard Disk Drive)",
                                            "4" => "SSD (Solid State Drive)",
                                            _ => "Unknown"
                                        };

                                        return driveType;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {

            }
            return "Error";
        }

        private static async Task<string?> GetDeviceIdFromDriveAsync(string driveRoot)
        {
            string driveLetter = driveRoot.TrimEnd('\\'); // Clean the input

            return await Task.Run(async () =>
            {
                string query = $"SELECT * FROM Win32_DiskDrive";
                using (var searcher = new ManagementObjectSearcher(query))
                {
                    foreach (ManagementObject logicalDisk in searcher.Get().Cast<ManagementObject>())
                    {
                        if ((string)logicalDisk["DeviceID"] == await GetDeviceIdFromDriveAsync(driveRoot))
                            return logicalDisk["DeviceID"]?.ToString(); // Return the DeviceID of the disk drive
                    }
                }
                return "Unknown"; // Fallback if not found
            });
        }

        private static string GetDriveTypeDescriptionAsync(string driveRoot)
        {
            Win32Helper.DriveType driveType = Win32Helper.GetDriveType(driveRoot);

            return driveType switch
            {
                Win32Helper.DriveType.DRIVE_REMOVABLE => "Removable",
                Win32Helper.DriveType.DRIVE_FIXED => GetDiskDriveFromLetter(driveRoot),
                Win32Helper.DriveType.DRIVE_REMOTE => "Network",
                Win32Helper.DriveType.DRIVE_CDROM => "CD-ROM",
                Win32Helper.DriveType.DRIVE_RAMDISK => "RAM Disk",
                Win32Helper.DriveType.DRIVE_NO_ROOT_DIR => "No Root Directory",
                _ => "Unknown",
            };
        }

        private void Button_Click(object sender, SplitButtonClickEventArgs e)
        {
            OptimizeSelected(AdvancedView.IsOn);
        }

        public static void RestartAsAdmin(string args)
        {
            var packageName = Package.Current.Id.FamilyName;
            var appId = CoreApplication.Id;

            // Request elevation
            var startInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-Command \"Start-Process 'shell:AppsFolder\\{packageName}!{appId}' -ArgumentList @('{args}') -Verb RunAs\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                Verb = "runas"
            };

            try
            {
                Process.Start(startInfo);
                App.Current.Exit();
            }
            catch
            {

            }
        }

        public async void OptimizeSelected(bool systemPartitions)
        {
            AdvancedView.IsOn = systemPartitions;

            if (!IsAdministrator())
            {
                RestartAsAdmin($"SELECTED{(systemPartitions ? "-SYSTEM" : "")} {MyListView.SelectedIndex}");
                return;
            }

            foreach (DiskItem item in MyListView.SelectedItems.Cast<DiskItem>())
            {
                string scriptPath = "C:\\Rebound11\\rdfrgui.ps1";
                string? volume = item.DriveLetter?.ToString().Remove(1, 2);
                string arguments = $@"
$job = Start-Job -ScriptBlock {{
    $global:OutputLines = @()
    
    # Capture all output including verbose messages
    Optimize-Volume -DriveLetter {volume} -Defrag -Verbose | ForEach-Object {{
        # Check if it's a progress update
        if ($_ -like ""*Progress*"") {{
            Write-Output ""Progress: $_""  # Modify to capture progress as needed
        }} else {{
            $global:OutputLines += $_
            Write-Output $_  # Output normal messages
        }}
    }}
}}

while ($job.State -eq 'Running') {{
    Clear-Host
    Start-Sleep -Seconds 0.01
    $output = Receive-Job -Id $job.Id -Keep
    if ($output) {{
        Write-Output $output[-1]  # Output the last line
    }}
}}

# Final output after the job completes
Receive-Job -Id $job.Id | ForEach-Object {{ Write-Output $_ }}

";

                // Ensure the script is written before proceeding
                await File.WriteAllTextAsync(scriptPath, arguments);

                try
                {
                    var processInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-ExecutionPolicy Bypass -File \"{scriptPath}\"",  // Use -File to execute the script
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        Verb = "runas"  // Run as administrator
                    };

                    using var process = new Process { StartInfo = processInfo };

                    string outputData = "0";
                    bool updateData = true;

                    var alreadyUsed = new List<string>();

                    process.OutputDataReceived += UpdateOutput;

                    void UpdateOutput(object sender, DataReceivedEventArgs args)
                    {
                        // Only process if there's data
                        if (!string.IsNullOrEmpty(args.Data))
                        {
                            // Use the dispatcher to update the UI
                            DispatcherQueue.TryEnqueue(async () => { await UpdateIO(args.Data); });

                            // Store the output data
                            outputData = "\n" + args.Data;
                        }
                    }

                    process.Start();
                    process.BeginOutputReadLine();

                    LoadSelectedItemInfo("Optimizing...");
                    //CurrentProgress.IsIndeterminate = true;
                    OptimizeButton.IsEnabled = false;
                    AnalyzeButton.IsEnabled = false;
                    VisualStateManager.GoToState(OptimizeButton, "Disabled", true);
                    CurrentDisk.Visibility = Visibility.Visible;
                    CurrentDisk.Text = "Optimizing...";
                    //UpdateIO();

                    async Task UpdateIO(string data)
                    {
                        if (!updateData) return;

                        if (data.Contains("VERBOSE: ") && data.Contains(" complete."))
                        {
                            string a = data.Substring(data.LastIndexOf("VERBOSE: ")).Replace("VERBOSE: ", string.Empty);

                            string dataToReplace = " complete.";

                            DispatcherQueue.TryEnqueue(() => { RunUpdate(a, dataToReplace); });
                        }

                        //UpdateIO();
                    }

                    void RunUpdate(string a, string dataToReplace)
                    {
                        if (alreadyUsed.Contains(a) != true)
                        {
                            alreadyUsed.Add(a);
                            CurrentProgress.Value = GetMaxPercentage(a);
                            if (a.Contains(" complete..."))
                            {
                                dataToReplace = " complete...";
                                if ((item as DiskItem).DriveLetter.ToString().Contains(@"}") != true) CurrentDisk.Text = $"Drive {volume}: - {a.Remove(a.IndexOf(" complete..."))}";
                                else CurrentDisk.Text = $"{(MyListView.SelectedItem as DiskItem).Name} - {a.Remove(a.IndexOf(" complete..."))}";
                            }
                            else if (a.Contains(" complete."))
                            {
                                dataToReplace = " complete.";
                                if ((item as DiskItem).DriveLetter.ToString().Contains(@"}") != true) CurrentDisk.Text = $"Drive {volume}: - {a.Remove(a.IndexOf(" complete."))}";
                                else CurrentDisk.Text = $"{(MyListView.SelectedItem as DiskItem).Name} - {a.Remove(a.IndexOf(" complete."))}";

                            }
                        }
                    }

                    await process.WaitForExitAsync();

                    updateData = false;
                    CurrentProgress.IsIndeterminate = false;
                    OptimizeButton.IsEnabled = true;
                    AnalyzeButton.IsEnabled = true;
                    CurrentDisk.Visibility = Visibility.Collapsed;
                    LoadSelectedItemInfo(GetStatus());

                    File.Delete(scriptPath);
                    alreadyUsed.Clear();
                    CurrentProgress.Value = 0;
                }
                catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 1223)
                {
                    ShowMessage("Defragmentation was canceled by the user.");
                }
                catch (Exception ex)
                {
                    ShowMessage($"Error: {ex.Message}");
                }
            }
        }

        static List<int> FindAllOccurrences(string mainString, string substring)
        {
            List<int> indices = new List<int>();
            int index = mainString.IndexOf(substring);

            while (index != -1)
            {
                indices.Add(index);
                index = mainString.IndexOf(substring, index + substring.Length);
            }

            return indices;
        }

        private int GetMaxPercentage(string data)
        {
            static string KeepOnlyNumbers(string input)
            {
                StringBuilder sb = new StringBuilder();

                foreach (char c in input)
                {
                    if (char.IsDigit(c))
                    {
                        sb.Append(c);
                    }
                }

                return sb.ToString();
            }
            var match = KeepOnlyNumbers(data);
            return int.Parse(match);
        }

        public async void OptimizeAll(bool close, bool systemPartitions)
        {
            AdvancedView.IsOn = systemPartitions;

            MyListView.IsEnabled = false;

            if (!IsAdministrator())
            {
                if (close == true) RestartAsAdmin($"OPTIMIZEALLANDCLOSE{(systemPartitions ? "-SYSTEM" : "")}");
                else RestartAsAdmin($"OPTIMIZEALL{(systemPartitions ? "-SYSTEM" : "")}");
                return;
            }

            int i = 0;
            int j = (MyListView.ItemsSource as List<DiskItem>).Count;

            MyListView.SelectedIndex = 0;

            foreach (var item in MyListView.ItemsSource as List<DiskItem>)
            {

                string scriptPath = "C:\\Rebound11\\rdfrgui.ps1";
                string volume = item.DriveLetter.ToString().Remove(1, 2);
                string arguments = $@"
$job = Start-Job -ScriptBlock {{
    $global:OutputLines = @()
    
    # Capture all output including verbose messages
    Optimize-Volume -DriveLetter {volume} -Defrag -Verbose | ForEach-Object {{
        # Check if it's a progress update
        if ($_ -like ""*Progress*"") {{
            Write-Output ""Progress: $_""  # Modify to capture progress as needed
        }} else {{
            $global:OutputLines += $_
            Write-Output $_  # Output normal messages
        }}
    }}
}}

while ($job.State -eq 'Running') {{
    Clear-Host
    Start-Sleep -Seconds 0.01
    $output = Receive-Job -Id $job.Id -Keep
    if ($output) {{
        Write-Output $output[-1]  # Output the last line
    }}
}}

# Final output after the job completes
Receive-Job -Id $job.Id | ForEach-Object {{ Write-Output $_ }}

";

                // Ensure the script is written before proceeding
                await File.WriteAllTextAsync(scriptPath, arguments);

                try
                {
                    i++;

                    if (DetailsBar.Severity == InfoBarSeverity.Error)
                    {
                        return;
                    }
                        var processInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-ExecutionPolicy Bypass -File \"{scriptPath}\"",  // Use -File to execute the script
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        Verb = "runas"  // Run as administrator
                    };

                    using var process = new Process { StartInfo = processInfo };

                    string outputData = "0";
                    bool updateData = true;

                    var alreadyUsed = new List<string>();

                    process.OutputDataReceived += UpdateOutput;

                    async void UpdateOutput(object sender, DataReceivedEventArgs args)
                    {
                        // Only process if there's data
                        if (!string.IsNullOrEmpty(args.Data))
                        {
                            // Use the dispatcher to update the UI
                            DispatcherQueue.TryEnqueue(async () => { await UpdateIO(args.Data); });

                            // Store the output data
                            outputData = "\n" + args.Data;
                        }
                    }

                    process.Start();
                    process.BeginOutputReadLine();

                    LoadSelectedItemInfo("Optimizing...");
                    //CurrentProgress.IsIndeterminate = true;
                    OptimizeButton.IsEnabled = false;
                    AnalyzeButton.IsEnabled = false;
                    VisualStateManager.GoToState(OptimizeButton, "Disabled", true);
                    CurrentDisk.Visibility = Visibility.Visible;
                    CurrentDisk.Text = "Optimizing...";
                    //UpdateIO();

                    async Task UpdateIO(string data)
                    {
                        if (!updateData) return;

                        if (data.Contains("VERBOSE: ") && data.Contains(" complete."))
                        {
                            string a = data.Substring(data.LastIndexOf("VERBOSE: ")).Replace("VERBOSE: ", string.Empty);

                            string dataToReplace = " complete.";

                            DispatcherQueue.TryEnqueue(() => { RunUpdate(a, dataToReplace); });
                        }

                        //UpdateIO();
                    }

                    async void RunUpdate(string a, string dataToReplace)
                    {
                        if (alreadyUsed.Contains(a) != true)
                        {
                            alreadyUsed.Add(a);
                            CurrentProgress.Value = GetMaxPercentage(a);
                            if (a.Contains(" complete..."))
                            {
                                dataToReplace = " complete...";
                                if ((item as DiskItem).DriveLetter.ToString().Contains(@"}") != true) CurrentDisk.Text = $"Drive {i}/{j} ({volume}:) - {a.Remove(a.IndexOf(" complete..."))}";
                                else CurrentDisk.Text = $"{(MyListView.SelectedItem as DiskItem).Name} ({i}/{j}) - {a.Remove(a.IndexOf(" complete..."))}";
                            }
                            else if (a.Contains(" complete."))
                            {
                                dataToReplace = " complete.";
                                if ((item as DiskItem).DriveLetter.ToString().Contains(@"}") != true) CurrentDisk.Text = $"Drive {i}/{j} ({volume}:) - {a.Remove(a.IndexOf(" complete."))}";
                                else CurrentDisk.Text = $"{(MyListView.SelectedItem as DiskItem).Name} ({i}/{j}) - {a.Remove(a.IndexOf(" complete."))}";

                            }
                        }
                    }

                    await process.WaitForExitAsync();

                    updateData = false;
                    CurrentProgress.IsIndeterminate = false;
                    OptimizeButton.IsEnabled = true;
                    AnalyzeButton.IsEnabled = true;
                    CurrentDisk.Visibility = Visibility.Collapsed;
                    LoadSelectedItemInfo(GetStatus());

                    File.Delete(scriptPath);
                    alreadyUsed.Clear();
                    CurrentProgress.Value = 0;
                    if (MyListView.SelectedIndex + 1 != j) MyListView.SelectedIndex++;
                }
                catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 1223)
                {
                    ShowMessage("Defragmentation was canceled by the user.");
                }
                catch (Exception ex)
                {
                    ShowMessage($"Error: {ex.Message}");
                }




                /*





                string volume = (item as DiskItem).DriveLetter.ToString().Remove(1, 2);
                string arguments = $"Optimize-Volume -DriveLetter {volume}"; // /O to optimize the drive
                                                                             //string arguments = $"Defrag {volume}: /O /U"; // /O to optimize the drive

                if (!IsAdmin())
                {
                    ShowMessage("The application must be running as Administrator to perform this task.");
                    return;
                }

                try
                {
                    i++;

                    var processInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = arguments,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        Verb = "runas"  // Run as administrator
                    };

                    var process = new Process
                    {
                        StartInfo = processInfo,
                        //EnableRaisingEvents = true
                    };

                    //process.OutputDataReceived += (s, ea) => UpdateProgress(ea.Data);
                    //process.ErrorDataReceived += (s, ea) => UpdateProgress(ea.Data);

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    if (DetailsBar.Severity == InfoBarSeverity.Error)
                    {
                        process.Close();
                        LoadSelectedItemInfo("Skipping...");
                        CurrentProgress.IsIndeterminate = true;
                        OptimizeButton.IsEnabled = false;
                        VisualStateManager.GoToState(OptimizeButton, "Disabled", true);
                        CurrentDisk.Visibility = Visibility.Visible;
                        if ((item as DiskItem).DriveLetter.ToString().Contains(@"}") != true) CurrentDisk.Text = $"Drive {i}/{j} ({volume}:) - Skipping...";
                        else CurrentDisk.Text = $"{(MyListView.SelectedItem as DiskItem).Name} ({i}/{j}) - Skipping...";

                        await Task.Delay(50);

                        LoadSelectedItemInfo(GetStatus());
                        CurrentProgress.IsIndeterminate = false;
                        OptimizeButton.IsEnabled = true;
                        CurrentDisk.Visibility = Visibility.Collapsed;
                    }
                    else
                    {
                        LoadSelectedItemInfo("Optimizing...");
                        CurrentProgress.IsIndeterminate = true;
                        OptimizeButton.IsEnabled = false;
                        VisualStateManager.GoToState(OptimizeButton, "Disabled", true);
                        CurrentDisk.Visibility = Visibility.Visible;
                        if ((item as DiskItem).DriveLetter.ToString().Contains(@"}") != true) CurrentDisk.Text = $"Drive {i}/{j} ({volume}:) - Optimizing...";
                        else CurrentDisk.Text = $"{(MyListView.SelectedItem as DiskItem).Name} ({i}/{j}) - Optimizing...";

                        await process.WaitForExitAsync();

                        LoadSelectedItemInfo(GetStatus());
                        CurrentProgress.IsIndeterminate = false;
                        OptimizeButton.IsEnabled = true;
                        CurrentDisk.Visibility = Visibility.Collapsed;
                    }

                    if (MyListView.SelectedIndex + 1 != j) MyListView.SelectedIndex++;
                }
                catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 1223)
                {
                    // Error 1223 indicates that the operation was canceled by the user (UAC prompt declined)
                    ShowMessage("Defragmentation was canceled by the user.");
                }
                catch (Exception ex)
                {
                    ShowMessage($"Error: {ex.Message}");
                }*/
            }

            i = 0;

            MyListView.IsEnabled = true;

            if (close == true) Close();
        }

        private void UpdateProgress(string data)
        {
            if (string.IsNullOrWhiteSpace(data)) return;

            // Parse the progress data from PowerShell output
            var match = Regex.Match(data, @"\[(o+)(\s*)\]");

            if (match.Success)
            {
                int progress = (int)((double)match.Groups[1].Value.Length / (match.Groups[1].Value.Length + match.Groups[2].Value.Length) * 100);

                // Update the UI with the progress
                DispatcherQueue.TryEnqueue(() =>
                {
                    UpdateProgress2($"{progress}% completed");
                });
            }
            else
            {
                // Handle other types of messages
                DispatcherQueue.TryEnqueue(() =>
                {
                    UpdateProgress2(data + Environment.NewLine);
                });
            }
        }

        private void ParseOutput(string data)
        {
            // Example line to match: "50% complete."
            //UpdateProgress(data);
            if (data != null)
            {
                Debug.WriteLine(data);
                    int i = 0;
                    foreach (char c in data)
                    {
                        if (c == 'o') i++;
                    }
                    UpdateProgress($"{i}%");
                //ShowMessage(data);
            }
        }

        public void UpdateProgress2(string progress)
        {
            var selectedItem = new DiskItem();
            selectedItem.DriveLetter = "C:/";
            selectedItem = MyListView.SelectedItem as DiskItem;
            // Update UI from the main thread
            _ = DispatcherQueue.TryEnqueue(() =>
            {
                DetailsBar.Title = selectedItem.Name;
                string selI;
                try
                {
                    List<string> i = DefragInfo.GetEventLogEntriesForID(258);
                    selI = i.Last(s => s.Contains($"({selectedItem.DriveLetter.ToString().Remove(2, 1)})"));
                }
                catch
                {
                    selI = "Never....";
                }
                DetailsBar.Message = $"Media type: {selectedItem.MediaType}\nLast analyzed or optimized: {selI.Substring(0, selI.Length - 4)}\nCurrent status: Optimizing ({progress}%)";
            });
        }

        private async void ShowMessage(string message)
        {
            try
            {
                var dialog = new ContentDialog
                {
                    Title = "Defragmentation",
                    Content = message,
                    CloseButtonText = "OK",
                };

                dialog.XamlRoot = this.Content.XamlRoot;
                _ = await dialog.ShowAsync();
            }
            catch (Exception ex)
            {

            }
        }

        private void MenuFlyoutItem_Click(object sender, RoutedEventArgs e)
        {
            OptimizeSelected(AdvancedView.IsOn);
        }

        private void MenuFlyoutItem_Click_1(object sender, RoutedEventArgs e)
        {
            OptimizeAll(false, AdvancedView.IsOn);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            //CreateModalWindow(this, new TaskWindow(), true, true);
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            LoadSelectedItemInfo(GetStatus());
        }

        private async void AdvancedView_Toggled(object sender, RoutedEventArgs e)
        {
            await LoadData(AdvancedView.IsOn);
        }
    }
    public class DefragInfo
    {
        public static DateTime? GetLastDefragTime(string driveLetter)
        {
            string driveRoot = driveLetter + ":\\";
            string query = "*[System[Provider[@Name='Microsoft-Windows-Defrag'] and (EventID=258 or EventID=259)]]";

            EventLogQuery eventsQuery = new EventLogQuery("Application", PathType.LogName, query);
            eventsQuery.ReverseDirection = true;
            eventsQuery.TolerateQueryErrors = true;

            List<EventRecord> events = new List<EventRecord>();

            using (EventLogReader logReader = new EventLogReader(eventsQuery))
            {
                for (EventRecord eventInstance = logReader.ReadEvent(); eventInstance != null; eventInstance = logReader.ReadEvent())
                {
                    if (eventInstance.Properties.Count > 0)
                    {
                        string message = eventInstance.Properties.Last().Value.ToString();
                        if (message.Contains(driveRoot))
                        {
                            return eventInstance.TimeCreated?.ToLocalTime();
                        }
                    }
                }
            }

            return null;
        }

        public static List<string> GetEventLogEntriesForID(int eventID)
        {
            List<string> eventMessages = new List<string>();

            // Define the query
            string logName = "Application"; // Windows Logs > Application
            string queryStr = "*[System/EventID=" + eventID + "]";

            EventLogQuery query = new EventLogQuery(logName, PathType.LogName, queryStr);

            // Create the reader
            using (EventLogReader reader = new EventLogReader(query))
            {
                for (EventRecord eventInstance = reader.ReadEvent(); eventInstance != null; eventInstance = reader.ReadEvent())
                {
                    // Extract the message from the event
                    string sb = eventInstance.TimeCreated.ToString() + eventInstance.FormatDescription().ToString().Substring(eventInstance.FormatDescription().ToString().Length - 4);

                    eventMessages.Add(sb.ToString());
                }
            }

            return eventMessages;
        }
    }

    public class StorageDevice
    {
        public string DeviceID { get; set; }
        public string MediaType { get; set; }
    }

    public class StorageDeviceService
    {
        // For CreateFile to get handle to drive
        private const uint GENERIC_READ = 0x80000000;
        private const uint GENERIC_WRITE = 0x40000000;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint OPEN_EXISTING = 3;
        private const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;

        // CreateFile to get handle to drive
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern SafeFileHandle CreateFileW(
            [MarshalAs(UnmanagedType.LPWStr)]
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        // For control codes
        private const uint FILE_DEVICE_MASS_STORAGE = 0x0000002d;
        private const uint IOCTL_STORAGE_BASE = FILE_DEVICE_MASS_STORAGE;
        private const uint FILE_DEVICE_CONTROLLER = 0x00000004;
        private const uint IOCTL_SCSI_BASE = FILE_DEVICE_CONTROLLER;
        private const uint METHOD_BUFFERED = 0;
        private const uint FILE_ANY_ACCESS = 0;
        private const uint FILE_READ_ACCESS = 0x00000001;
        private const uint FILE_WRITE_ACCESS = 0x00000002;

        private static uint CTL_CODE(uint DeviceType, uint Function,
                                     uint Method, uint Access)
        {
            return ((DeviceType << 16) | (Access << 14) |
                    (Function << 2) | Method);
        }

        // For DeviceIoControl to check no seek penalty
        private const uint StorageDeviceSeekPenaltyProperty = 7;
        private const uint PropertyStandardQuery = 0;

        [StructLayout(LayoutKind.Sequential)]
        private struct STORAGE_PROPERTY_QUERY
        {
            public uint PropertyId;
            public uint QueryType;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
            public byte[] AdditionalParameters;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DEVICE_SEEK_PENALTY_DESCRIPTOR
        {
            public uint Version;
            public uint Size;
            [MarshalAs(UnmanagedType.U1)]
            public bool IncursSeekPenalty;
        }

        // DeviceIoControl to check no seek penalty
        [DllImport("kernel32.dll", EntryPoint = "DeviceIoControl",
                   SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeviceIoControl(
            SafeFileHandle hDevice,
            uint dwIoControlCode,
            ref STORAGE_PROPERTY_QUERY lpInBuffer,
            uint nInBufferSize,
            ref DEVICE_SEEK_PENALTY_DESCRIPTOR lpOutBuffer,
            uint nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped);

        // For DeviceIoControl to check nominal media rotation rate
        private const uint ATA_FLAGS_DATA_IN = 0x02;

        [StructLayout(LayoutKind.Sequential)]
        private struct ATA_PASS_THROUGH_EX
        {
            public ushort Length;
            public ushort AtaFlags;
            public byte PathId;
            public byte TargetId;
            public byte Lun;
            public byte ReservedAsUchar;
            public uint DataTransferLength;
            public uint TimeOutValue;
            public uint ReservedAsUlong;
            public IntPtr DataBufferOffset;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
            public byte[] PreviousTaskFile;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
            public byte[] CurrentTaskFile;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ATAIdentifyDeviceQuery
        {
            public ATA_PASS_THROUGH_EX header;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
            public ushort[] data;
        }

        // DeviceIoControl to check nominal media rotation rate
        [DllImport("kernel32.dll", EntryPoint = "DeviceIoControl",
                   SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeviceIoControl(
            SafeFileHandle hDevice,
            uint dwIoControlCode,
            ref ATAIdentifyDeviceQuery lpInBuffer,
            uint nInBufferSize,
            ref ATAIdentifyDeviceQuery lpOutBuffer,
            uint nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped);

        // For error message
        private const uint FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern uint FormatMessage(
            uint dwFlags,
            IntPtr lpSource,
            uint dwMessageId,
            uint dwLanguageId,
            StringBuilder lpBuffer,
            uint nSize,
            IntPtr Arguments);

        // Method for no seek penalty
        public static bool HasNoSeekPenalty(string sDrive)
        {
            SafeFileHandle hDrive = CreateFileW(
                sDrive,
                0, // No access to drive
                FILE_SHARE_READ | FILE_SHARE_WRITE,
                IntPtr.Zero,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                IntPtr.Zero);

            if (hDrive == null || hDrive.IsInvalid)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                Debug.WriteLine("CreateFile failed. " + message);
            }

            uint IOCTL_STORAGE_QUERY_PROPERTY = CTL_CODE(
                IOCTL_STORAGE_BASE, 0x500,
                METHOD_BUFFERED, FILE_ANY_ACCESS); // From winioctl.h

            STORAGE_PROPERTY_QUERY query_seek_penalty =
                new STORAGE_PROPERTY_QUERY();
            query_seek_penalty.PropertyId = StorageDeviceSeekPenaltyProperty;
            query_seek_penalty.QueryType = PropertyStandardQuery;

            DEVICE_SEEK_PENALTY_DESCRIPTOR query_seek_penalty_desc =
                new DEVICE_SEEK_PENALTY_DESCRIPTOR();

            uint returned_query_seek_penalty_size;

            bool query_seek_penalty_result = DeviceIoControl(
                hDrive,
                IOCTL_STORAGE_QUERY_PROPERTY,
                ref query_seek_penalty,
                (uint)Marshal.SizeOf(query_seek_penalty),
                ref query_seek_penalty_desc,
                (uint)Marshal.SizeOf(query_seek_penalty_desc),
                out returned_query_seek_penalty_size,
                IntPtr.Zero);

            hDrive.Close();

            if (query_seek_penalty_result == false)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                Debug.WriteLine("DeviceIoControl failed. " + message);
                return false;
            }
            else
            {
                if (query_seek_penalty_desc.IncursSeekPenalty == false)
                {
                    Debug.WriteLine("This drive has NO SEEK penalty.");
                    return false;
                }
                else
                {
                    Debug.WriteLine("This drive has SEEK penalty.");
                    return true;
                }
            }
        }

        // Method for nominal media rotation rate
        // (Administrative privilege is required)
        public static bool HasNominalMediaRotationRate(string sDrive)
        {
            SafeFileHandle hDrive = CreateFileW(
                sDrive,
                GENERIC_READ | GENERIC_WRITE, // Administrative privilege is required
                FILE_SHARE_READ | FILE_SHARE_WRITE,
                IntPtr.Zero,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                IntPtr.Zero);

            if (hDrive == null || hDrive.IsInvalid)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                Debug.WriteLine("CreateFile failed. " + message);
            }

            uint IOCTL_ATA_PASS_THROUGH = CTL_CODE(
                IOCTL_SCSI_BASE, 0x040b, METHOD_BUFFERED,
                FILE_READ_ACCESS | FILE_WRITE_ACCESS); // From ntddscsi.h

            ATAIdentifyDeviceQuery id_query = new ATAIdentifyDeviceQuery();
            id_query.data = new ushort[256];

            id_query.header.Length = (ushort)Marshal.SizeOf(id_query.header);
            id_query.header.AtaFlags = (ushort)ATA_FLAGS_DATA_IN;
            id_query.header.DataTransferLength =
                (uint)(id_query.data.Length * 2); // Size of "data" in bytes
            id_query.header.TimeOutValue = 3; // Sec
            id_query.header.DataBufferOffset = (IntPtr)Marshal.OffsetOf(
                typeof(ATAIdentifyDeviceQuery), "data");
            id_query.header.PreviousTaskFile = new byte[8];
            id_query.header.CurrentTaskFile = new byte[8];
            id_query.header.CurrentTaskFile[6] = 0xec; // ATA IDENTIFY DEVICE

            uint retval_size;

            bool result = DeviceIoControl(
                hDrive,
                IOCTL_ATA_PASS_THROUGH,
                ref id_query,
                (uint)Marshal.SizeOf(id_query),
                ref id_query,
                (uint)Marshal.SizeOf(id_query),
                out retval_size,
                IntPtr.Zero);

            hDrive.Close();

            if (result == false)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                Debug.WriteLine("DeviceIoControl failed. " + message);
                return true;
            }
            else
            {
                // Word index of nominal media rotation rate
                // (1 means non-rotate device)
                const int kNominalMediaRotRateWordIndex = 217;

                if (id_query.data[kNominalMediaRotRateWordIndex] == 1)
                {
                    Debug.WriteLine("This drive is NON-ROTATE device.");
                    return false;
                }
                else
                {
                    Debug.WriteLine("This drive is ROTATE device.");
                    return true;
                }
            }
        }

        // Method for error message
        private static string GetErrorMessage(int code)
        {
            StringBuilder message = new StringBuilder(255);

            FormatMessage(
              FORMAT_MESSAGE_FROM_SYSTEM,
              IntPtr.Zero,
              (uint)code,
              0,
              message,
              (uint)message.Capacity,
              IntPtr.Zero);

            return message.ToString();
        }

        // For control codes
        private const uint IOCTL_VOLUME_BASE = 0x00000056;

        // For DeviceIoControl to get disk extents
        [StructLayout(LayoutKind.Sequential)]
        private struct DISK_EXTENT
        {
            public uint DiskNumber;
            public long StartingOffset;
            public long ExtentLength;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct VOLUME_DISK_EXTENTS
        {
            public uint NumberOfDiskExtents;
            [MarshalAs(UnmanagedType.ByValArray)]
            public DISK_EXTENT[] Extents;
        }

        // DeviceIoControl to get disk extents
        [DllImport("kernel32.dll", EntryPoint = "DeviceIoControl",
                   SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeviceIoControl(
            SafeFileHandle hDevice,
            uint dwIoControlCode,
            IntPtr lpInBuffer,
            uint nInBufferSize,
            ref VOLUME_DISK_EXTENTS lpOutBuffer,
            uint nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped);

        // Method for disk extents
        public static int GetDiskExtents(char cDrive)
        {
            DriveInfo di = new DriveInfo(cDrive.ToString());
            if (di.DriveType != DriveType.Fixed)
            {
                Debug.WriteLine("This drive is not fixed drive.");
            }

            string sDrive = "\\\\.\\" + cDrive.ToString() + ":";

            SafeFileHandle hDrive = CreateFileW(
                sDrive,
                0, // No access to drive
                FILE_SHARE_READ | FILE_SHARE_WRITE,
                IntPtr.Zero,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                IntPtr.Zero);

            if (hDrive == null || hDrive.IsInvalid)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                Debug.WriteLine("CreateFile failed. " + message);
            }

            uint IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS = CTL_CODE(
                IOCTL_VOLUME_BASE, 0,
                METHOD_BUFFERED, FILE_ANY_ACCESS); // From winioctl.h

            VOLUME_DISK_EXTENTS query_disk_extents =
                new VOLUME_DISK_EXTENTS();

            uint returned_query_disk_extents_size;

            bool query_disk_extents_result = DeviceIoControl(
                hDrive,
                IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS,
                IntPtr.Zero,
                0,
                ref query_disk_extents,
                (uint)Marshal.SizeOf(query_disk_extents),
                out returned_query_disk_extents_size,
                IntPtr.Zero);

            hDrive.Close();

            if (query_disk_extents_result == false ||
                query_disk_extents.Extents.Length != 1)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                Debug.WriteLine("DeviceIoControl failed. " + message);
            }
            else
            {
                Debug.WriteLine("The physical drive number is: " +
                                  query_disk_extents.Extents[0].DiskNumber);
            }

            return (int)query_disk_extents.Extents[0].DiskNumber;
        }
    }
    public class VolumeInfo
    {
        public string GUID { get; set; }
        public string FileSystem { get; set; }
        public ulong Size { get; set; }
        public string FriendlyName { get; set; }
    }

    public class SystemVolumes
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool GetVolumeInformation(
            string rootPathName,
            StringBuilder volumeNameBuffer,
            int volumeNameSize,
            out uint volumeSerialNumber,
            out uint maximumComponentLength,
            out uint fileSystemFlags,
            StringBuilder fileSystemNameBuffer,
            int nFileSystemNameSize);

        public static List<VolumeInfo> GetSystemVolumes()
        {
            List<VolumeInfo> volumes = new List<VolumeInfo>();

            // WMI query to get all volumes, including GUID paths
            string query = "SELECT * FROM Win32_Volume WHERE DriveLetter IS NULL";
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(query))
            {
                foreach (ManagementObject volume in searcher.Get())
                {
                    string volumePath = volume["DeviceID"].ToString(); // This gives the \\?\Volume{GUID} path
                    string fileSystem = volume["FileSystem"]?.ToString() ?? "Unknown";
                    ulong size = (ulong)volume["Capacity"];

                    // We can further refine this by querying for EFI, Recovery, etc., based on size and file system
                    string friendlyName;
                    if (fileSystem == "FAT32" && size < 512 * 1024 * 1024)
                    {
                        friendlyName = "EFI System Partition";
                    }
                    else if (fileSystem == "NTFS" && size > 500 * 1024 * 1024)
                    {
                        friendlyName = "Recovery Partition";
                    }
                    else if (fileSystem == "NTFS" && size < 500 * 1024 * 1024)
                    {
                        friendlyName = "System Reserved Partition";
                    }
                    else
                    {
                        friendlyName = "Unknown System Partition";
                    }

                    volumes.Add(new VolumeInfo
                    {
                        GUID = volumePath,
                        FileSystem = fileSystem,
                        Size = size,
                        FriendlyName = friendlyName
                    });
                }
            }

            return volumes;
        }
    }

    public class SSDOrHDDDriveInfo
    {
        public string Identifier { get; }
        public string Model { get; private set; }
        public string Type { get; private set; }

        public SSDOrHDDDriveInfo(string identifier)
        {
            Identifier = identifier;
            GetDriveType();
        }

        private void GetDriveType()
        {
            string query = string.IsNullOrWhiteSpace(Identifier) ?
                "SELECT * FROM Win32_DiskDrive" :
                $"SELECT * FROM Win32_DiskDrive WHERE DeviceID='{Identifier}' OR DeviceID='{GetVolumeGUID(Identifier)}'";

            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(query))
            {
                foreach (ManagementObject drive in searcher.Get())
                {
                    Model = drive["Model"].ToString();
                    string mediaType = drive["MediaType"]?.ToString() ?? "Unknown";
                    Type = mediaType.Contains("SSD") ? "SSD" : "HDD";
                    return; // Exit after the first matching drive
                }
            }

            // If no drive found, set default values
            Model = "Unknown";
            Type = "Unknown";
        }

        private static string GetVolumeGUID(string driveLetter)
        {
            // Convert drive letter to GUID format if necessary
            return $"\\\\?\\Volume{{{driveLetter.TrimEnd('\\').ToUpper()}}}\\";
        }
    }
}
