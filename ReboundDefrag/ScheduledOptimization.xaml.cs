using Microsoft.UI.Xaml.Media;
using ReboundDefrag.Helpers;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;
using WinUIEx;
using static ReboundDefrag.MainWindow;

namespace ReboundDefrag
{
    public sealed partial class ScheduledOptimization : WindowEx
    {
        public ScheduledOptimization(int parentX, int parentY)
        {
            this.InitializeComponent();
            Win32Helper.RemoveIcon(this);
            IsMaximizable = false;
            IsMinimizable = false;
            this.MoveAndResize(parentX + 50, parentY + 50, 550, 600);
            IsResizable = false;
            Title = "Scheduled optimization";
            SystemBackdrop = new MicaBackdrop();
            AppWindow.DefaultTitleBarShouldMatchAppModeTheme = true;
            LoadData();
        }

        public async void LoadData()
        {
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
                            item.IsChecked = GetBoolFromLocalSettings(newDriveLetter);
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

            int selIndex = MyListView.SelectedIndex is not -1 ? MyListView.SelectedIndex : 0;

            // Set the list view's item source
            MyListView.ItemsSource = items;

            MyListView.SelectedIndex = selIndex >= items.Count ? items.Count - 1 : selIndex;
        }
    }
}
