using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Text;

#nullable enable
#pragma warning disable SYSLIB1054 // Use 'LibraryImportAttribute' instead of 'DllImportAttribute' to generate P/Invoke marshalling code at compile time
#pragma warning disable CA1401 // P/Invokes should not be visible

namespace ReboundDefrag.Helpers
{
    public static partial class Win32Helper
    {
        public const int WM_DEVICECHANGE = 0x0219;
        public const int DBT_DEVICEARRIVAL = 0x8000;
        public const int DBT_DEVICEREMOVECOMPLETE = 0x8004;

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint GetLogicalDrives();

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool GetVolumeInformation(
            string lpRootPathName,
            StringBuilder lpVolumeNameBuffer,
            int nVolumeNameSize,
            out uint lpVolumeSerialNumber,
            out uint lpMaximumComponentLength,
            out uint lpFileSystemFlags,
            StringBuilder lpFileSystemNameBuffer,
            int nFileSystemNameSize);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern DriveType GetDriveType(string lpRootPathName);

        public enum DriveType : uint
        {
            DRIVE_UNKNOWN = 0,
            DRIVE_NO_ROOT_DIR = 1,
            DRIVE_REMOVABLE = 2,
            DRIVE_FIXED = 3,
            DRIVE_REMOTE = 4,
            DRIVE_CDROM = 5,
            DRIVE_RAMDISK = 6
        }
    }
}
