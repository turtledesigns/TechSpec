using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Web.Script.Serialization;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Xml.Linq;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using System.Text;

namespace TechSpec
{
    public class TechSpec_Main
    {
        // For CreateFile to get handle to drive
        public const uint GENERIC_READ = 0x80000000;
        public const uint GENERIC_WRITE = 0x40000000;
        public const uint FILE_SHARE_READ = 0x00000001;
        public const uint FILE_SHARE_WRITE = 0x00000002;
        public const uint OPEN_EXISTING = 3;
        public const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;

        // CreateFile to get handle to drive
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern SafeFileHandle CreateFileW(
            [MarshalAs(UnmanagedType.LPWStr)]
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        // For control codes
        public const uint FILE_DEVICE_MASS_STORAGE = 0x0000002d;
        public const uint IOCTL_STORAGE_BASE = FILE_DEVICE_MASS_STORAGE;
        public const uint FILE_DEVICE_CONTROLLER = 0x00000004;
        public const uint IOCTL_SCSI_BASE = FILE_DEVICE_CONTROLLER;
        public const uint METHOD_BUFFERED = 0;
        public const uint FILE_ANY_ACCESS = 0;
        public const uint FILE_READ_ACCESS = 0x00000001;
        public const uint FILE_WRITE_ACCESS = 0x00000002;

        public static uint CTL_CODE(uint DeviceType, uint Function,
                                     uint Method, uint Access)
        {
            return ((DeviceType << 16) | (Access << 14) |
                    (Function << 2) | Method);
        }

        // For DeviceIoControl to check no seek penalty
        public const uint StorageDeviceSeekPenaltyProperty = 7;
        public const uint PropertyStandardQuery = 0;

        [StructLayout(LayoutKind.Sequential)]
        public struct STORAGE_PROPERTY_QUERY
        {
            public uint PropertyId;
            public uint QueryType;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
            public byte[] AdditionalParameters;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct DEVICE_SEEK_PENALTY_DESCRIPTOR
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
        public static extern bool DeviceIoControl(
            SafeFileHandle hDevice,
            uint dwIoControlCode,
            ref STORAGE_PROPERTY_QUERY lpInBuffer,
            uint nInBufferSize,
            ref DEVICE_SEEK_PENALTY_DESCRIPTOR lpOutBuffer,
            uint nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped);

        // For DeviceIoControl to check nominal media rotation rate
        public const uint ATA_FLAGS_DATA_IN = 0x02;

        [StructLayout(LayoutKind.Sequential)]
        public struct ATA_PASS_THROUGH_EX
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
        public struct ATAIdentifyDeviceQuery
        {
            public ATA_PASS_THROUGH_EX header;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
            public ushort[] data;
        }

        // DeviceIoControl to check nominal media rotation rate
        [DllImport("kernel32.dll", EntryPoint = "DeviceIoControl",
                   SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DeviceIoControl(
            SafeFileHandle hDevice,
            uint dwIoControlCode,
            ref ATAIdentifyDeviceQuery lpInBuffer,
            uint nInBufferSize,
            ref ATAIdentifyDeviceQuery lpOutBuffer,
            uint nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped);

        // For error message
        public const uint FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern uint FormatMessage(
            uint dwFlags,
            IntPtr lpSource,
            uint dwMessageId,
            uint dwLanguageId,
            StringBuilder lpBuffer,
            uint nSize,
            IntPtr Arguments);

        static void Main(string[] args)
        {
            HasNominalMediaRotationRate("ERROR TEST");
        }


        // Method for nominal media rotation rate
        // (Administrative privilege is required)
        public static string HasNominalMediaRotationRate(string name)
        {
            SafeFileHandle hDrive = CreateFileW(
                name,
                GENERIC_READ | GENERIC_WRITE, // Administrative privilege is required
                FILE_SHARE_READ | FILE_SHARE_WRITE,
                IntPtr.Zero,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                IntPtr.Zero);

            if (hDrive == null || hDrive.IsInvalid)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                return ("CreateFile failed. " + message);
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
                return ("DeviceIoControl failed. " + message);
            }
            else
            {
                // Word index of nominal media rotation rate
                // (1 means non-rotate device)
                const int kNominalMediaRotRateWordIndex = 217;

                if (id_query.data[kNominalMediaRotRateWordIndex] == 1)
                {
                    return "true";
                }
                else
                {
                    return "Disk is non SSD";
                }

            }
        }

        // Method for error message
        public static string GetErrorMessage(int code)
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

        public static int writeFileToJSON(string uri)
        {
            try
            {
                // create the path if it doesn't exist
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(uri));

                System.IO.StreamWriter file = new System.IO.StreamWriter(uri);
                file.WriteLine(getJSONString());
                file.Close();
                return 1;
            }
            catch
            {
                return 0;
            }
        }

        public static string getJSONString()
        {
            return (new JavaScriptSerializer().Serialize(getJSONObject()));
        }

        private static object getJSONObject()
        {
            object obj = new Lad
            {
                cpu = new CPU
                {
                    manufacturer = GetCPUManufacturer(),
                    model = GetCPU(),
                    speed = GetCPUSpeed(),
                },
                model = GetModel(),
                os = new OperatingSystem
                {
                    name = GetOS(),
                    version = GetOSVersion()
                },
                manufacturer = GetManufacturer(),
                screensize = GetScreenSize(),
                storage = GetStorage(),
                opticaldrive = GetOpticalDrive(),
                ram = GetRAM(),
                graphics = GetGraphics(),
                screenResolution = new ScreenResolution
                {
                    width = GetScreenWidth(),
                    height = GetScreenHeight()
                },
                sound = GetSound(),
                formFactor = GetFormFactor(),
                wireless = GetWireless(),
                touchscreen = GetTouchscreen(),
                webcam = GetWebcam(),
                bluetooth = GetBluetooth(),
                macAddress = GetMACAddress()
            };

            return obj;
        }

        public static string GetManufacturer()
        {
            string multipleVariables = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables += wmi.GetPropertyValue("Manufacturer").ToString();
                }
                catch { }
            }
            if (multipleVariables == "")
            {
                return "N/A";
            }
            else
            {
                return multipleVariables;
            }
        }

        public static string GetScreenManufacturer()
        {
            string multipleVariables = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DesktopMonitor");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables += wmi.GetPropertyValue("Caption").ToString();
                }
                catch { }
            }
            if (multipleVariables == "")
            {
                return "N/A";
            }
            else
            {
                return multipleVariables;
            }
        }

        public static string GetModel()
        {
            string multipleVariables = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_Computersystem");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables += wmi.GetPropertyValue("Model").ToString();
                }
                catch { }
            }
            if (multipleVariables == "")
            {
                return "N/A";
            }
            else
            {
                return multipleVariables;
            }
        }

        public static string GetFormFactor()
        {
            string multipleVariables = "";
            int formFactor = 0;

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_Computersystem");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    formFactor = Convert.ToInt32(wmi.GetPropertyValue("PCSystemType"));
                    if (formFactor == 2)
                    {
                        multipleVariables = "Mobile";
                    } else
                    {
                        multipleVariables = "Desktop";
                    }
                }
                catch { }
            }
            if (multipleVariables == "")
            {
                return "N/A";
            }
            else
            {
                return multipleVariables;
            }
        }

        public static string GetOS()
        {
            string multipleVariables = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_OperatingSystem");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables += wmi.GetPropertyValue("Caption").ToString();
                }
                catch { }
            }
            if (multipleVariables == "")
            {
                return "N/A";
            }
            else
            {
                return multipleVariables;
            }
        }

        public static string GetOSVersion()
        {
            string multipleVariables = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_OperatingSystem");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables += wmi.GetPropertyValue("Version").ToString();
                }
                catch { }
            }
            if (multipleVariables == "")
            {
                return "N/A";
            }
            else
            {
                return multipleVariables;
            }
        }

        public static string GetScreenWidth()
        {
            string multipleVariables = "0";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DesktopMonitor");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables = wmi.GetPropertyValue("ScreenWidth").ToString();
                }
                catch { }
            }
            return multipleVariables;
        }

        public static string GetScreenHeight()
        {
            string multipleVariables = "0";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DesktopMonitor");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables = wmi.GetPropertyValue("ScreenHeight").ToString();
                }
                catch { }
            }
            return multipleVariables;
        }

        public static string GetMACAddress()
        {
            string multipleVariables = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_NetworkAdapterConfiguration");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables = wmi.GetPropertyValue("MacAddress").ToString();
                }
                catch { }
            }
            if (multipleVariables == "")
            {
                return "N/A";
            }
            else
            {
                return multipleVariables;
            }
        }

        static private string RunDxDiag()
        {
            var psi = new ProcessStartInfo();
            if (IntPtr.Size == 4 && Environment.Is64BitOperatingSystem)
            {
                psi.FileName = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    "sysnative\\dxdiag.exe");
            }
            else
            {
                psi.FileName = System.IO.Path.Combine(
                    Environment.SystemDirectory,
                    "dxdiag.exe");
            }
            string path = System.IO.Path.GetTempFileName();
            try
            {
                psi.Arguments = "/x " + path;
                using (var prc = Process.Start(psi))
                {
                    prc.WaitForExit();
                    if (prc.ExitCode != 0)
                    {
                        throw new Exception("DXDIAG failed with exit code " + prc.ExitCode.ToString());
                    }
                }
                return System.IO.File.ReadAllText(path);
            }
            finally
            {
                System.IO.File.Delete(path);
            }
        }

        public static IDictionary<string, GraphicsDriveList> GetGraphics()
        {
            IDictionary<string, GraphicsDriveList> graphicsObject = new Dictionary<string, GraphicsDriveList>();
            int counter = 0;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM CIM_VideoController");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    if (wmi.GetPropertyValue("InstalledDisplayDrivers").ToString() != "")
                    {
                        List<string> values = convertToMBGBTB(Convert.ToDouble(wmi.GetPropertyValue("AdapterRAM"))).Split(',').ToList<string>();

                        int fourGBCheck = Convert.ToInt32(values[0]);
                        if (fourGBCheck == 4)
                        {
                            string DedicatedMemory = "";
                            string modelName = wmi.GetPropertyValue("VideoProcessor").ToString();
                            string xml = RunDxDiag();

                            XElement root = XElement.Parse(xml);


                            IEnumerable<XElement> purchaseOrders =
                                from el in root.Element("DisplayDevices").Elements("DisplayDevice")
                                where
                                    (string)el.Element("ChipType") == modelName
                                select el;
                            foreach (XElement el in purchaseOrders)
                                DedicatedMemory = (string)el.Element("DedicatedMemory");

                            if (DedicatedMemory.IndexOf(" MB") != -1)
                            {
                                double valueInMB = Convert.ToDouble(DedicatedMemory.Remove(DedicatedMemory.Length - 3));
                                DedicatedMemory = Math.Round((valueInMB / 1024), 0) + " GB";
                            }
                            else if (DedicatedMemory.IndexOf(" GB") != -1)
                            {
                                values[1] = "GB";
                            }
                            else if (DedicatedMemory.IndexOf(" TB") != -1)
                            {
                                values[1] = "TB";
                            }
                            else
                            {
                                values[1] = DedicatedMemory + "****";
                            }
                            DedicatedMemory = DedicatedMemory.Remove(DedicatedMemory.Length - 3);
                            values[0] = DedicatedMemory;
                        };



                        graphicsObject[counter.ToString()] = new GraphicsDriveList
                        {
                            manufacturer = wmi.GetPropertyValue("AdapterCompatibility").ToString(),
                            model = wmi.GetPropertyValue("VideoProcessor").ToString(),
                            adapterRamSize = Convert.ToInt32(values[0]),
                            adapterRamValue = values[1],
                            resolutionWidth = Convert.ToInt32(wmi.GetPropertyValue("CurrentHorizontalResolution")),
                            resolutionHeight = Convert.ToInt32(wmi.GetPropertyValue("CurrentVerticalResolution")),
                        };

                        counter++;
                    }
                }
                catch { }

            }

            if (counter > 0)
            {
                return graphicsObject;
            }
            else
            {
                graphicsObject["0"] = new GraphicsDriveList
                {
                    manufacturer = "",
                    model = "",
                    adapterRamSize = 0,
                    adapterRamValue = "",
                    resolutionWidth = 0,
                    resolutionHeight = 0,
                };

                return graphicsObject;
            }

        }

        /**
         * Returns a list of graphics card manufacturers
         * 
         */
        public static List<string> GetGraphicsManufacturers()
        {
            // Get all the graphics card information
            IDictionary<string, GraphicsDriveList> graphics = GetGraphics();

            // Extract the manufacturer and add to a list
            List<string> manufacturers = new List<string>();
            foreach (KeyValuePair<string, GraphicsDriveList> entry in graphics)
            {
                manufacturers.Add(entry.Value.manufacturer);
            }

            return manufacturers;
        }

        private static IDictionary<string, SoundDevices> GetSound()
        {
            IDictionary<string, SoundDevices> soundObject = new Dictionary<string, SoundDevices>();
            int counter = 0;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_SoundDevice");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    if (wmi.GetPropertyValue("Manufacturer").ToString() != "")
                    {
                        soundObject[counter.ToString()] = new SoundDevices
                        {
                            manufacturer = wmi.GetPropertyValue("Manufacturer").ToString(),
                            name = wmi.GetPropertyValue("Caption").ToString(),
                        };
                        counter++;
                    }
                }
                catch { }

            }

            if (counter > 0)
            {
                return soundObject;
            }
            else
            {
                return null;
            }

        }

        public static IDictionary<string, ScreenSizes> GetScreenSize()
        {
            IDictionary<int, ScreenSizeLookup> screenSizeLookup = new Dictionary<int, ScreenSizeLookup>();
            
            screenSizeLookup[0] = new ScreenSizeLookup { cm = 33, inches = 13, };
            screenSizeLookup[1] = new ScreenSizeLookup { cm = 34, inches = 13.3, };
            screenSizeLookup[2] = new ScreenSizeLookup { cm = 35, inches = 13.9, };
            screenSizeLookup[3] = new ScreenSizeLookup { cm = 36, inches = 14, };
            screenSizeLookup[4] = new ScreenSizeLookup { cm = 40, inches = 15.6, };
            screenSizeLookup[5] = new ScreenSizeLookup { cm = 44, inches = 17.3, };
            screenSizeLookup[6] = new ScreenSizeLookup { cm = 47, inches = 18.4, };            

            IDictionary<string, ScreenSizes> screenSize = new Dictionary<string, ScreenSizes>();
            int counter = 0;
            double screenWidth = 0;
            double screenHeight = 0;
            double screenSizeInCM = 0;
            double screenSizeInInches = 0;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("\\root\\wmi", "SELECT MaxVerticalImageSize,MaxHorizontalImageSize FROM WmiMonitorBasicDisplayParams");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    screenWidth = Convert.ToDouble(wmi.GetPropertyValue("MaxVerticalImageSize"));
                    screenHeight = Convert.ToDouble(wmi.GetPropertyValue("MaxHorizontalImageSize"));
                    screenSizeInCM = Convert.ToDouble(Math.Sqrt((screenHeight * screenHeight) + (screenWidth * screenWidth)));

                    if (GetFormFactor() == "Mobile")
                    {
                        double closestMatchingSizeValue = -999;
                        for (int i = screenSizeLookup.Count - 1; i >= 0; i--)
                        {
                            var item = screenSizeLookup.ElementAt(i);
                            var itemValueInCM = item.Value.cm;
                            var itemValueInInches = item.Value.inches;

                            var temp = item.Value.cm - screenSizeInCM;
                            if (temp < 0)
                            {
                                temp = temp * -1;
                            }

                            if (closestMatchingSizeValue == -999 || temp < closestMatchingSizeValue)
                            {
                                closestMatchingSizeValue = temp;
                                screenSizeInInches = item.Value.inches;
                            }
                        }
                    } else
                    {
                        screenSizeInInches = (Math.Round(screenSizeInCM / 2.54, 0));
                    }


                    screenSize[counter.ToString()] = new ScreenSizes
                    {
                        size = screenSizeInInches.ToString(),
                    };

                    counter++;
                }
                catch { }

            }

            if (counter > 0)
            {
                return screenSize;
            }
            else
            {
                return null;
            }
        }

        private static string convertToMBGBTB(double value)
        {
            string extension = ",MB";

            if (value != 0)
            {
                value = value / 1048576;

                if (value >= 1024)
                {
                    value = value / 1024;
                    extension = ",GB";
                }
                if (value >= 1024)
                {
                    value = value / 1024;
                    extension = ",TB";
                }
            }

            if (value != 0)
            {
                return Math.Round(value, 0) + extension;
            }
            else
            {
                return "N/A";
            }
        }

        private static string convertToMBGBTBHardDrive(double value)
        {
            string extension = ",KB";

            if (value != 0)
            {
                value = value / 1000;

                if (value >= 1000)
                {
                    value = value / 1000;
                    extension = ",MB";
                }
                if (value >= 1000)
                {
                    value = value / 1000;
                    extension = ",GB";
                }
                if (value >= 1000)
                {
                    value = value / 1000;
                    extension = ",TB";
                }
            }

            if (value != 0)
            {
                return Math.Round(value, 0) + extension;
            }
            else
            {
                return "N/A";
            }
        }

        public static string GetCPU()
        {
            string multipleVariables = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_Processor");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables = Regex.Replace(wmi.GetPropertyValue("Name").ToString(), @"\s+", " ");
                    multipleVariables = Regex.Replace(multipleVariables, @"\(R\)", "®");
                    multipleVariables = Regex.Replace(multipleVariables, @"\(r\)", "®");
                    multipleVariables = Regex.Replace(multipleVariables, @"\(TM\)", "™");
                    multipleVariables = Regex.Replace(multipleVariables, @"\(tm\)", "™");
                    multipleVariables = Regex.Replace(multipleVariables, @"Genuine\s", "");
                    multipleVariables = Regex.Replace(multipleVariables, @"(.*) CPU", "$1");
                    multipleVariables = Regex.Replace(multipleVariables, @"(.*) @ (.*)GHz", "$1");
                    multipleVariables = Regex.Replace(multipleVariables, @"(.*)(\s\S*)-Core Processor", "$1");
                }
                catch { }
            }
            if (multipleVariables == "")
            {
                return "N/A";
            }
            else
            {
                return multipleVariables;
            }
        }

        public static string GetCPUManufacturer()
        {
            string multipleVariables = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_Processor");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables = wmi.GetPropertyValue("Manufacturer").ToString();
                }
                catch { }
            }
            if (multipleVariables == "")
            {
                return "N/A";
            }
            else
            {
                return multipleVariables;
            }
        }

        public static string GetCPUSpeed()
        {
            string multipleVariables = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_Processor");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables = (Convert.ToDouble(wmi.GetPropertyValue("MaxClockSpeed")) / 1000).ToString("0.00") +"GHz";
                }
                catch { }
            }
            if (multipleVariables == "")
            {
                return "N/A";
            }
            else
            {
                return multipleVariables;
            }
        }

        private static IDictionary<string, StorageDrives> GetStorage()
        {
            double storageSize = 0;
            int counter = 0;
            
            IDictionary<string, StorageDrives> storageObject = new Dictionary<string, StorageDrives>();
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT Caption,Size,Name FROM Win32_DiskDrive WHERE MediaType = 'Fixed hard disk media'");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    if ((wmi.GetPropertyValue("Caption").ToString().IndexOf("SSD") != -1) || HasNominalMediaRotationRate(wmi.GetPropertyValue("Name").ToString()) == "true")
                    {
                        storageSize = Convert.ToDouble(wmi.GetPropertyValue("Size"));

                        if (storageSize != 0)
                        {
                            List<string> values = convertToMBGBTBHardDrive(storageSize).Split(',').ToList<string>();

                            storageObject[counter.ToString()] = new StorageDrives
                            {
                                size = Convert.ToInt32(values[0]),
                                value = values[1],
                                type = "SSD",
                                caption = wmi.GetPropertyValue("Caption").ToString()
                            };

                            counter++;
                        }
                    }
                    else
                    {
                        storageSize = Convert.ToDouble(wmi.GetPropertyValue("Size"));

                        if (storageSize != 0)
                        {
                            List<string> values = convertToMBGBTBHardDrive(storageSize).Split(',').ToList<string>();

                            storageObject[counter.ToString()] = new StorageDrives
                            {
                                size = Convert.ToInt32(values[0]),
                                value = values[1],
                                type = "HDD",
                                caption = wmi.GetPropertyValue("Caption").ToString()
                            };

                            counter++;
                        }
                    }

                }
                catch { }
            }

            if (counter > 0)
            {
                return storageObject;
            }
            else
            {
                storageObject["0"] = new StorageDrives
                {
                    size = 0,
                    value = "N/A",
                    type = "N/A",
                    caption = "N/A"
                };

                return storageObject;
            }
        }

        private static RAM GetRAM()
        {
            double ramSpeed = 0;
            double ramCapacity = 0;
            int memoryValue = 0;
            string memoryType = "";
            int SMBIOS_MemoryValue = 0;
            string SMBIOS_MemoryType = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PhysicalMemory WHERE FormFactor > 0");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    memoryValue = Convert.ToInt32(wmi.GetPropertyValue("MemoryType"));
                    SMBIOS_MemoryValue = Convert.ToInt32(wmi.GetPropertyValue("SMBIOSMemoryType"));
                    double ramValues = Convert.ToDouble(wmi.GetPropertyValue("Capacity"));
                    ramCapacity += ramValues;

                    if (ramSpeed <= Convert.ToDouble(wmi.GetPropertyValue("Speed"))) {
                        ramSpeed = Convert.ToDouble(wmi.GetPropertyValue("Speed"));
                    }

                }
                catch { }
            }

            switch (memoryValue)
            {
                case 1:
                    memoryType = "";
                    break;
                case 2:
                    memoryType = "DRAM";
                    break;
                case 3:
                    memoryType = "Synchronous DRAM";
                    break;
                case 4:
                    memoryType = "Cache DRAM";
                    break;
                case 5:
                    memoryType = "EDO";
                    break;
                case 6:
                    memoryType = "EDRAM";
                    break;
                case 7:
                    memoryType = "VRAM";
                    break;
                case 8:
                    memoryType = "SRAM";
                    break;
                case 9:
                    memoryType = "RAM";
                    break;
                case 10:
                    memoryType = "ROM";
                    break;
                case 11:
                    memoryType = "Flash";
                    break;
                case 12:
                    memoryType = "EEPROM";
                    break;
                case 13:
                    memoryType = "FEPROM";
                    break;
                case 14:
                    memoryType = "EPROM";
                    break;
                case 15:
                    memoryType = "CDRAM";
                    break;
                case 16:
                    memoryType = "3DRAM";
                    break;
                case 17:
                    memoryType = "SDRAM";
                    break;
                case 18:
                    memoryType = "SGRAM";
                    break;
                case 19:
                    memoryType = "RDRAM";
                    break;
                case 20:
                    memoryType = "DDR";
                    break;
                case 21:
                    memoryType = "DDR2";
                    break;
                case 22:
                    memoryType = "FB-DDR2";
                    break;
                default:
                    if (memoryValue > 22)
                        memoryType = "DDR3";
                    else
                        memoryType = "";
                    break;
            }

            switch (SMBIOS_MemoryValue)
            {
                case 12:
                    SMBIOS_MemoryType = "EEPROM";
                    break;
                case 13:
                    SMBIOS_MemoryType = "CDRAM";
                    break;
                case 14:
                    SMBIOS_MemoryType = "3DRAM";
                    break;
                case 15:
                    SMBIOS_MemoryType = "SDRAM";
                    break;
                case 16:
                    SMBIOS_MemoryType = "SGRAM";
                    break;
                case 17:
                    SMBIOS_MemoryType = "RDRAM";
                    break;
                case 18:
                    SMBIOS_MemoryType = "DDR";
                    break;
                case 19:
                    SMBIOS_MemoryType = "DDR2";
                    break;
                case 20:
                    SMBIOS_MemoryType = "FB-DDR2";
                    break;
                case 21:
                    SMBIOS_MemoryType = "Reserved";
                    break;
                case 22:
                    SMBIOS_MemoryType = "Reserved";
                    break;
                case 23:
                    SMBIOS_MemoryType = "Reserved";
                    break;
                case 24:
                    SMBIOS_MemoryType = "DDR3";
                    break;
                case 25:
                    SMBIOS_MemoryType = "FBD2";
                    break;
                case 26:
                    SMBIOS_MemoryType = "DDR4";
                    break;
                case 27:
                    SMBIOS_MemoryType = "IPDDR";
                    break;
                case 28:
                    SMBIOS_MemoryType = "IPDDR2";
                    break;
                case 29:
                    SMBIOS_MemoryType = "IPDDR3";
                    break;
                case 30:
                    SMBIOS_MemoryType = "IPDDR4";
                    break;
                default:
                    SMBIOS_MemoryType = "";
                    break;
            }




            if (ramCapacity == 0 || ramSpeed == 0)
            {
                return new RAM
                {
                    size = 0,
                    value = "N/A",
                    type = "N/A",
                    SMBIOS_Type = "N/A",
                    speed = 0
                };
            }
            else
            {
                List<string> values = convertToMBGBTB(ramCapacity).Split(',').ToList<string>();

                return new RAM
                {
                    size = Convert.ToInt32(values[0]),
                    value = values[1],
                    type = memoryType,
                    SMBIOS_Type = SMBIOS_MemoryType,
                    speed = ramSpeed
                };
            }


        }

        public static int GetWireless()
        {
            int wirelessAvailable = 0;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE DHCPEnabled = 'TRUE'");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    string wireless = wmi.GetPropertyValue("Description").ToString();

                    if (!(wireless.IndexOf("Wi-Fi") == -1 && wireless.IndexOf("Wireless") == -1 && wireless.IndexOf("wifi") == -1))
                    {
                        wirelessAvailable = 1;
                    }
                }
                catch { }
            }

            return wirelessAvailable;
        }

        public static string GetOpticalDrive()
        {
            string multipleVariables = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_CDRomDrive");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables += wmi.GetPropertyValue("MediaType").ToString();
                }
                catch { }
            }
            if (multipleVariables == "")
            {
                return "N/A";
            }
            else
            {
                return multipleVariables;
            }
        }

        public static int GetTouchscreen()
        {
            int multipleVariables = 0;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PnpEntity  WHERE Description LIKE '%touch screen%'");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables = (wmi.GetPropertyValue("Description").ToString() != "") ? 1 : 0;
                }
                catch { }
            }

            return multipleVariables;
        }

        public static int GetWebcam()
        {
            int multipleVariables = 0;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PnpEntity  WHERE Caption LIKE '%cam%'");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables = (wmi.GetPropertyValue("Caption").ToString() != "") ? 1 : 0;
                }
                catch { }
            }

            return multipleVariables;
        }

        public static int GetBluetooth()
        {
            int multipleVariables = 0;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PnpEntity  WHERE Description LIKE '%bluetooth%'");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    multipleVariables = (wmi.GetPropertyValue("Description").ToString() != "") ? 1 : 0;
                }
                catch { }
            }

            return multipleVariables;
        }

        private class Lad
        {
            public CPU cpu;
            public string model;
            public OperatingSystem os;
            public string manufacturer;
            public IDictionary<string, ScreenSizes> screensize;
            public IDictionary<string, SoundDevices> sound;
            public IDictionary<string, StorageDrives> storage;
            public string opticaldrive;
            public RAM ram;
            public IDictionary<string, GraphicsDriveList> graphics;
            public ScreenResolution screenResolution;
            public string formFactor;
            public int wireless;
            public int touchscreen;
            public int webcam;
            public int bluetooth;
            public string macAddress;
        }

        private class OperatingSystem
        {
            public string name;
            public string version;
        }

        private class ScreenResolution
        {
            public string width;
            public string height;
        }

        private class StorageDrives
        {
            public int size;
            public string value;
            public string type;
            public string caption;
        }

        private class SoundDevices
        {
            public string manufacturer;
            public string name;
        }

        public class GraphicsDriveList
        {
            public string manufacturer;
            public string model;
            public int adapterRamSize;
            public string adapterRamValue;
            public int resolutionWidth;
            public int resolutionHeight;
        }

        public class ScreenSizes
        {
            public string size;
        }

        private class CPU
        {
            public string manufacturer;
            public string model;
            public string speed;
        }

        private class RAM
        {
            public int size;
            public string value;
            public string type;
            public string SMBIOS_Type;
            public double speed;
        }

        private class ScreenSizeLookup
        {
            public double cm;
            public double inches;
        }


    }

}