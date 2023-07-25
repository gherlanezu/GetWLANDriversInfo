///-----------------------------------------------------------------
///   Namespace:      GetWLANDriversInfo
///   Class:          Program.cs
///   Description:    Main routines for the IT scheduled task for wireless jobs
///   Author:         Dan Codorean                     Date: 2012
///   Notes:          Configures the scheduled task to run it once a day, copy/updates the local wireless utilities when conencted to corp network.
///                   
///   Revision History: I will start the revision history in April 2022, when I finally check in the code in GitHub
///   Name:           Date:        Description:
///-----------------------------------------------------------------
///

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Runtime.InteropServices; // dllimport
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Security.Principal;
using System.Threading;
using System.Text;
using System.Xml;
using Microsoft.Win32; // for registry
//using TaskScheduler;


namespace GetWLANDriversInfo
{

    public enum RunningPlatform
    {
        _32 = 0,
        _64,
        unknown
    }

    public enum WLANClientUpdateStatus
    {
        CURRENT = 0, // 
        UPDATE_AVAILABLE,
        NOT_APPLICABLE,
        CANNOT_READ_XML_CONFIG,
        UNKNOWN
    }

    public enum PowerShellExecutionPolicy
    {
        Restricted,
        AllSigned,
        RemoteSigned,
        Unrestricted
    }

    public enum NetworkContext
    {
        insideIntel = 0,
        insideIntelWLAN,
        insideIntelWired,
        insideIntelVPN,
        outsideIntel,
        unknown
    }

    public enum FileUpdateResult
    {
        successful = 0,
        identical_files,
        src_not_found,
        failed_to_copy,
        not_on_corp_network,
        access_denied
    }

    public class ThreadParameters
    {
        private string mvar_sCmdLine;
        private string mvar_sArgs;
        private ProcessWindowStyle mvar_winStyle;
        private bool mvar_bWaitForThread;
        private long mvar_lReturnCode;

        public string CMDLine
        {
            get { return mvar_sCmdLine; }
            set { this.mvar_sCmdLine = value; }
        }

        public string CMDLineArguments
        {
            get { return mvar_sArgs; }
            set { this.mvar_sArgs = value; }
        }

        public bool WaitForThread
        {
            get { return mvar_bWaitForThread; }
            set { this.mvar_bWaitForThread = value; }
        }

        public ProcessWindowStyle ProcWindowStyle
        {
            get { return mvar_winStyle; }
            set { this.mvar_winStyle = value; }
        }

        public long ReturnCode
        {
            get { return mvar_lReturnCode; }
            set { this.mvar_lReturnCode = value; }
        }

        public ThreadParameters()
        {
            mvar_winStyle = ProcessWindowStyle.Normal;
            mvar_sCmdLine = string.Empty;
            mvar_sArgs = string.Empty;
            mvar_bWaitForThread = false;
            mvar_lReturnCode = -1;
        }

        public ThreadParameters(string cmdLine, string arguments, ProcessWindowStyle winStyle, bool waitForIt)
        {
            mvar_bWaitForThread = waitForIt;
            mvar_winStyle = winStyle;
            mvar_sCmdLine = cmdLine;
            mvar_sArgs = arguments;
            mvar_lReturnCode = -1;
        }
    }

    public class iLogFile
    {
        string mvar_sFileName;
        string mvar_sFilePath;
        string mvar_sFullName;
        bool mvar_bLogLocked;

        public iLogFile()
        {
            mvar_sFileName = "GetITWLANInfo.log"; //System.Windows.Forms.Application.ProductName + System.Windows.Forms.Application.ProductVersion + ".log";
            mvar_sFilePath = Environment.GetEnvironmentVariable("SystemDrive") + "\\Intel\\Logs";
            mvar_sFullName = mvar_sFilePath + "\\" + mvar_sFileName;
            mvar_bLogLocked = false;
        }

        public iLogFile(iLogFile val)
        {
            mvar_sFileName = val.logFile;
            mvar_sFilePath = val.logPath;
            mvar_sFullName = val.logFullName;
            mvar_bLogLocked = false;
        }

        public iLogFile(string fileName)
        {
            if (fileName.Contains("\\"))
            {
                mvar_sFullName = fileName;
                string[] aTemp = fileName.Split('\\');
                int pos = aTemp.Length - 1;
                mvar_sFileName = aTemp[pos];
                mvar_sFilePath = fileName.Replace("\\" + mvar_sFileName, "");
            }
            else
            {
                mvar_sFileName = fileName;
                mvar_sFilePath = Environment.GetEnvironmentVariable("SystemDrive") + "\\Intel\\Logs";
                mvar_sFullName = mvar_sFilePath + "\\" + mvar_sFileName;
            }
            mvar_bLogLocked = false;
        }

        public iLogFile(string fileName, string filePath)
        {
            mvar_sFileName = fileName;
            mvar_sFilePath = filePath;
            mvar_sFullName = filePath + "\\" + fileName;
            mvar_bLogLocked = false;
        }

        ~iLogFile()
        { }


        public string logFile
        {
            get { return mvar_sFileName; }
            set { this.mvar_sFileName = value.Length > 0 ? value : ""; }
        }

        public string logPath
        {
            get { return mvar_sFilePath; }
            set { this.mvar_sFilePath = value.Length > 0 ? value : ""; }
        }

        public string logFullName
        {
            get { return mvar_sFullName; }
            set { this.mvar_sFullName = value.Length > 0 ? value : ""; }
        }

        public bool IsLocked
        {
            get { return mvar_bLogLocked; }
            set { this.mvar_bLogLocked = value; }
        }

        public void LogInfo(string info2Log)
        {
            if (!Directory.Exists(mvar_sFilePath))
                Directory.CreateDirectory(mvar_sFilePath);

            try
            {
                if (File.Exists(mvar_sFullName))
                {
                    FileInfo fi = new FileInfo(mvar_sFullName);
                    // if the log file is larger then 512k (half meg), copy it into the hist log and crete a new log.
                    if (fi.Length > 512000)
                    {
                        File.Copy(mvar_sFullName, GetNextBackupFile(mvar_sFullName), true);
                        File.Delete(mvar_sFullName);
                    }
                }
                int cnt = 0;
                while (IsLocked)
                {
                    Thread.Sleep(200);
                    //System.Windows.Forms.Application.DoEvents();
                    cnt += 1;
                    // in case the log is locked by another thread.. wait.. but only for one second ... 
                    // if not done in 1 second attempt to write to it anyway...
                    if (cnt > 5)
                        break;
                }

                using (StreamWriter log = File.AppendText(mvar_sFullName))
                {
                    IsLocked = true;
                    lock (log)
                    {
                        log.WriteLine(info2Log);
                        Debug.Print(info2Log);
                        log.Flush();
                        log.Close();
                    }
                    IsLocked = false;
                    return;
                }
            }
            catch (System.Exception ex)
            {
                iLogFile errLog = new iLogFile();
                errLog.LogInfo("err while logging: " + ex.Message + "\r\n" + ex.StackTrace);
            }
        }

        public string GetNextBackupFile(string fileName)
        {
            string retValue = fileName + ".hist.log";
            FileInfo fi = new FileInfo(fileName);
            string fileFolder = fi.DirectoryName;
            string baseFileName = fi.Name.Replace(fi.Extension, "");
            string backupFile = fileFolder + "\\" + baseFileName + "hist";
            for (int i = 9; i > 0; i--)
            {
                if (File.Exists(backupFile + i.ToString() + fi.Extension))
                    File.Copy(backupFile + i + fi.Extension, backupFile + (i + 1) + fi.Extension, true);

            }
            retValue = backupFile + "1" + fi.Extension;
            return retValue;
        }

        public string LogHeader()
        {
            return "[" + DateTime.Now + "] ";
        }

        public void LogFull(string info2log)
        {
            LogInfo(LogHeader() + info2log);
        }

    }

    public class ChkItem
    {

        public enum itemType
        {
            unknown = 0,
            reg_val,
            reg_key,
            reg_val_subtree,
            file_ver,
            app_ver,
            drv_ver,
            wlandriver_ver,
            wmi,
            service,
            os_ver,
            sp_ver,
            ie_ver,
            exact_copy
        }

        public enum itemStatus
        {
            CURRENT = 0,
            NOT_CURRENT,
            UNDETERMINED
        }

        public enum comparisonOperatorType
        {
            equal = 0,
            not_equal,
            greater_then,
            greater_then_or_equal,
            less_then,
            less_then_or_equal,
            oneof,
            substring
        }

        private string m_name;
        private itemType m_type;
        private string m_provider;
        private string m_parameter;
        private string m_property;
        private short m_left;
        private short m_right;
        private string m_compareto;
        private string m_textIfTrue;
        private string m_textIfFalse;
        private comparisonOperatorType m_operator;
        private itemStatus m_status;

        public ChkItem(string name, itemType type, string scope, string provider, string parameter, string prop, short fromLeft, short fromRight, comparisonOperatorType compOperator,
        string compareto, itemStatus status, string textIfTrue, string textIfFalse)
        {
            m_name = name;
            m_type = type;
            m_provider = provider;
            m_parameter = parameter;
            m_property = prop;
            m_left = fromLeft;
            m_right = fromRight;
            m_operator = compOperator;
            m_compareto = compareto;
            m_status = status;
            m_textIfTrue = textIfTrue;
            m_textIfFalse = textIfFalse;
        }


        public ChkItem()
        {
            m_name = string.Empty;
            m_type = itemType.unknown;
            m_provider = string.Empty;
            m_parameter = string.Empty;
            m_property = string.Empty;
            m_left = 0;
            m_right = 0;
            m_operator = comparisonOperatorType.equal;
            m_compareto = string.Empty;
            m_status = itemStatus.UNDETERMINED;
            m_textIfTrue = string.Empty;
            m_textIfFalse = string.Empty;
        }

        public ChkItem(ChkItem src)
        {
            m_name = src.Name;
            m_type = src.Type;
            m_provider = src.Provider;
            m_parameter = src.Parameter;
            m_property = src.Property;
            m_left = src.Left;
            m_right = src.Right;
            m_operator = src.ComparisonOperator;
            m_compareto = src.CompareTO;
            m_status = src.Status;
            m_textIfTrue = src.TextIfTrue;
            m_textIfFalse = src.TextIfFalse;
        }

        public string Name
        {
            get { return m_name; }
            set { this.m_name = value; }
        }

        public itemType Type
        {
            get { return m_type; }
            set { this.m_type = value; }
        }

        public string Provider
        {
            get { return m_provider; }
            set { this.m_provider = value; }
        }

        public string Parameter
        {
            get { return m_parameter; }
            set { this.m_parameter = value; }
        }

        public string Property
        {
            get { return m_property; }
            set { this.m_property = value; }
        }

        public short Left
        {
            get { return m_left; }
            set { this.m_left = value; }
        }

        public short Right
        {
            get { return m_right; }
            set { this.m_right = value; }
        }

        public comparisonOperatorType ComparisonOperator
        {
            get { return m_operator; }
            set { this.m_operator = value; }
        }

        public string CompareTO
        {
            get { return m_compareto; }
            set { this.m_compareto = value; }
        }

        public itemStatus Status
        {
            get { return m_status; }
            set { this.m_status = value; }
        }

        public string TextIfTrue
        {
            get { return m_textIfTrue; }
            set
            {
                //this.m_textIfTrue = value;
                if (value.Length < 1 | (value == null))
                {
                    m_textIfTrue = string.Empty;
                }
                else
                {
                    m_textIfTrue = value;
                }
            }
        }

        public string TextIfFalse
        {
            get { return m_textIfFalse; }
            set { this.m_textIfFalse = value; }
        }

        public static itemType ItemTypeFromString(string itemTypeString)
        {
            itemType functionReturnValue = itemType.unknown;
            switch (itemTypeString.ToLower())
            {
                case "unknown":
                    functionReturnValue = itemType.unknown;
                    break;
                case "reg_val":
                    functionReturnValue = itemType.reg_val;
                    break;
                case "reg_key":
                    functionReturnValue = itemType.reg_key;
                    break;
                case "reg_val_subtree":
                    functionReturnValue = itemType.reg_val_subtree;
                    break;
                case "file_ver":
                    functionReturnValue = itemType.file_ver;
                    break;
                case "app_ver":
                    functionReturnValue = itemType.app_ver;
                    break;
                case "drv_ver":
                    functionReturnValue = itemType.drv_ver;
                    break;
                case "wlandriver_ver":
                    functionReturnValue = itemType.wlandriver_ver;
                    break;
                case "wmi":
                    functionReturnValue = itemType.wmi;
                    break;
                case "service":
                    functionReturnValue = itemType.service;
                    break;
                case "os_ver":
                    functionReturnValue = itemType.os_ver;
                    break;
                case "sp_ver":
                    functionReturnValue = itemType.sp_ver;
                    break;
                case "ie_ver":
                    functionReturnValue = itemType.ie_ver;
                    break;
                case "exact_copy":
                    functionReturnValue = itemType.exact_copy;
                    break;
                default:
                    functionReturnValue = itemType.unknown;
                    break;
            }
            return functionReturnValue;
        }

        public static comparisonOperatorType ComparisonOperatorFromString(string @operator)
        {
            comparisonOperatorType functionReturnvalue = comparisonOperatorType.equal;
            switch (@operator)
            {
                case "!=":
                    functionReturnvalue = comparisonOperatorType.not_equal;
                    break;
                case ">":
                case "gt":
                    functionReturnvalue = comparisonOperatorType.greater_then;
                    break;
                case ">=":
                case "gte":
                    functionReturnvalue = comparisonOperatorType.greater_then_or_equal;
                    break;
                case "<":
                case "lt":
                    functionReturnvalue = comparisonOperatorType.less_then;
                    break;
                case "<=":
                case "lte":
                    functionReturnvalue = comparisonOperatorType.less_then_or_equal;
                    break;
                default:
                    functionReturnvalue = comparisonOperatorType.equal;
                    break;
            }
            return functionReturnvalue;
        }

    }

    public class WirelessLANInfo
    {
        private string m_strAdapterName;
        private string m_strDeviceID;
        private string m_strAppVersionFound; // actual
        private string m_strAppVersionCurrentRelease; // App (proset) version required on the system
        private string m_strDriverVer; // actual
        private string m_strDriverVerCurrentRelease; // current production / to be 
        private ChkItem m_chkIntem;
        private string m_strPathToInstaller;
        private string m_strInstaller;
        private string m_strCmdArguments;
        private string m_strPassThruPath;

        public WirelessLANInfo()
        {
            m_strAdapterName = string.Empty;
            m_strDeviceID = string.Empty;
            m_strAppVersionFound = string.Empty;
            m_strAppVersionCurrentRelease = string.Empty;
            m_strDriverVer = string.Empty;
            m_strDriverVerCurrentRelease = string.Empty;
            m_chkIntem = new ChkItem();
            m_strPathToInstaller = string.Empty;
            m_strInstaller = string.Empty;
            m_strCmdArguments = string.Empty;
            m_strPassThruPath = string.Empty;
        }


        public WirelessLANInfo(WirelessLANInfo src)
        {
            m_strAdapterName = src.AdapterName;
            m_strDeviceID = src.DeviceID;
            m_strAppVersionFound = src.AppVersionFound;
            m_strAppVersionCurrentRelease = src.AppVersionCurrentRelease;
            m_strDriverVer = src.DriverVersionFound;
            m_strDriverVerCurrentRelease = src.DriverVersionCurrentRelease;
            m_chkIntem = src.AdapterCheckItem;
            m_strPathToInstaller = src.PathToInstaller;
            m_strInstaller = src.Installer;
            m_strCmdArguments = src.CmdArguments;
            m_strPassThruPath = src.PassThruPath;
        }

        public string Print()
        {
            string printOut = "Wireless LAN Info values collected:\r\n" +
                "=========================================\r\n" +
                "\r\nAdapter Name: " + m_strAdapterName +
                "\r\nDevice ID: " + m_strDeviceID +
                "\r\nApp Version found: " + m_strAppVersionFound +
                "\r\nApp Version Current release: " + m_strAppVersionCurrentRelease +
                "\r\nDriver Version: " + m_strDriverVer;
            if (m_strPathToInstaller.Length > 0)
                printOut += "\r\nUpdate path: " + m_strPathToInstaller;


            return printOut;
        }

        public string AdapterName
        {
            get { return m_strAdapterName; }
            set { this.m_strAdapterName = value; }
        }

        public string DeviceID
        {
            get { return m_strDeviceID; }
            set { this.m_strDeviceID = value; }
        }

        public ChkItem AdapterCheckItem
        {
            get { return m_chkIntem; }
            set { this.m_chkIntem = value; }
        }

        public string AppVersionFound
        {
            get { return m_strAppVersionFound; }
            set { this.m_strAppVersionFound = value; }
        }

        public string AppVersionCurrentRelease
        {
            get { return m_strAppVersionCurrentRelease; }
            set { this.m_strAppVersionCurrentRelease = value; }
        }

        public string DriverVersionFound
        {
            get { return m_strDriverVer; }
            set { this.m_strDriverVer = value; }
        }

        public string DriverVersionCurrentRelease
        {
            get { return m_strDriverVerCurrentRelease; }
            set { this.m_strDriverVerCurrentRelease = value; }
        }

        public string PathToInstaller
        {
            get { return m_strPathToInstaller; }
            set { this.m_strPathToInstaller = value; }
        }

        public string Installer
        {
            get { return m_strInstaller; }
            set { this.m_strInstaller = value; }
        }

        public string CmdArguments
        {
            get { return m_strCmdArguments; }
            set { this.m_strCmdArguments = value; }
        }

        public string PassThruPath
        {
            get { return m_strPassThruPath; }
            set { this.m_strPassThruPath = value; }
        }
    }

    public class Program
    {

        public const string intelsm_path = "\\\\amr.corp.intel.com\\iss\\IntelSM";
        public const string ISS_Path = "\\\\amr.corp.intel.com\\iss\\olwnprod";
        public const string ISSTest_Path = "\\\\amr.corp.intel.com\\iss\\isstest";
        public const string ISS_APPID = "";

        public static string friendlyProductName = "Get Wireless LAN Client Info";
        public static NetworkContext currentNetworkContext;
        public static bool bCheck4Updates = false;
        public static ThreadParameters runShellArguments;
        public static Dictionary<string, string> AllWirelessConnections = new Dictionary<string, string>();
        public static RunningPlatform platform;
        public static Version osVer;
        public static string strOSVer = string.Empty;
        public static string output = string.Empty;
        public static string interfaceName = string.Empty;
        public static string DeviceInstanceID = string.Empty;
        public static string wirelessAdapter = string.Empty;
        public static string wirelessDriverVersion = string.Empty;
        public static string wirelessDriverFile = string.Empty;
        public static string CurrentPROSetVersion = string.Empty;
        public static string CurrentDriverPackage = string.Empty;
        public static string RegkeyAdapterPath = string.Empty;
        public static string CurrentLogedOnUser = string.Empty;
        public static bool bVerboseLog = false;
        public static bool bSilentRun = false;
        public static bool bWait = false;
        public static bool bFroceTaskCreation = false;
        public static bool bIsIntelAdapter = true;
        public static bool bFullRun = false;
        public static bool bVACGreenLight = false;

        public static string inforesult = string.Empty;

        public static List<string> AllInstalledApps;

        public static iLogFile logFile;
        public static WirelessLANInfo wlaninfo;
        public static WLANClientUpdateStatus wirelessUpdateStatus;

        public static RegistryKey regKeyPackage;

        //public static System.IO.StreamReader output;

        static int Main(string[] args)
        {
            int retCode = 1;
            if (args.Contains<string>("/v"))
                bVerboseLog = true;
            if (args.Contains<string>("/V"))
                bVerboseLog = true;
            if (args.Contains<string>("/s"))
                bSilentRun = true;
            if (args.Contains<string>("/S"))
                bSilentRun = true;
            if (args.Contains<string>("/w"))
                bWait = true;
            if (args.Contains<string>("/W"))
                bWait = true;
            if (args.Contains<string>("/f"))
                bFroceTaskCreation = true;
            if (args.Contains<string>("/F"))
                bFroceTaskCreation = true;
            logFile = new iLogFile();

            if (File.Exists(logFile.logFullName))
            {
                FileInfo fi = new FileInfo(logFile.logFullName);
                // if the log file is larger then 256k, delete it... 
                if (fi.Length > 256000)
                {
                    //File.Copy(mvar_sFullName, GetNextBackupFile(mvar_sFullName), true);
                    File.Delete(logFile.logFullName);
                }
            }

            Log("====================================================");
            Log("====================================================");
            Log("");
            Log("Starting GetWLANDriversInfo v." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() + " (" + DateTime.Now + ")");
            Log("");
            Log(Environment.CommandLine);
            Log("");
            WindowsIdentity id = WindowsIdentity.GetCurrent();
            CurrentLogedOnUser = id.Name;

            Log("Running as " + CurrentLogedOnUser, true);
            //osVer = GetOSFullVersion();
            //osVer = new Version(GetOSFullVersionSTR());
            osVer = GetOSVersionFromRegistry();
            strOSVer = osVer.Major + "." + osVer.Minor;
            Log("OS Version: " + osVer.ToString());

            Program.platform = SizeOfIntPtr();
            //string runtimestring = DateTime.Now.ToLongDateString();
            //runtimestring = DateTime.Now.ToShortDateString();
            //runtimestring = DateTime.Now.ToString("R");


            //try
            //{
            //    Console.WriteLine("In try ");

            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("In catch with " + ex.Message);
            //}

            try
            {
                currentNetworkContext = NetworkContext.unknown;
                currentNetworkContext = GetCurrentNetworkContext();
                string destPath = Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN";


                if (!System.IO.Directory.Exists(destPath))
                    System.IO.Directory.CreateDirectory(destPath);
                string currentFile = System.Reflection.Assembly.GetExecutingAssembly().Location; // System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
                string fileName = System.IO.Path.GetFileName(Assembly.GetExecutingAssembly().CodeBase);
                if (!currentFile.ToLower().Equals(destPath.ToLower()))
                {
                    if (bVerboseLog)
                    {
                        Log("Copy: " + currentFile + " -> " + Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\" + fileName, true);
                        //Console.WriteLine("Copy: " + currentFile + " -> " + Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\" + fileName);
                    }
                    CopyFileSafe(currentFile, Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\" + fileName);
                }

                ConfigureTheScheduledTask();
                AllInstalledApps = GetInstalledApplications();
                AllInstalledApps.Sort();

                if (currentNetworkContext == NetworkContext.insideIntel)
                {
                    bCheck4Updates = true;
                }


                regKeyPackage = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (regKeyPackage == null)
                {
                    Registry.LocalMachine.CreateSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
                    regKeyPackage = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
                }
                string strRunCount2Three = regKeyPackage.GetValue("RunID", string.Empty).ToString();
                int iRunID = 1;
                if (!strRunCount2Three.Equals(string.Empty))
                    iRunID = Int32.Parse(strRunCount2Three);
                bFullRun = false;
                if (iRunID == 1)
                    bFullRun = true;

                iRunID++;
                if (iRunID > 4)
                    regKeyPackage.SetValue("RunID", 1);
                else
                    regKeyPackage.SetValue("RunID", iRunID);

                /// first check if the system is on corp network and do updates if necessary
                /// 
                if (bCheck4Updates)
                {
                    if (bVerboseLog) Log("Update local files...", true);
                    FileUpdateResult updateStatus = FileUpdateResult.identical_files;
                    updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\WLANInfoConfig\WLANInfoConfig.xml", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\WLANInfoConfig.xml");
                    if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                    //if (System.IO.File.Exists(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\WLANInfoConfig\WLANInfoConfig.xml"))
                    //    CopyFileSafe(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\WLANInfoConfig\WLANInfoConfig.xml", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\WLANInfoConfig.xml");

                    if (!System.IO.Directory.Exists(Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\QoS"))
                        System.IO.Directory.CreateDirectory(Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\QoS");

                    if (System.IO.File.Exists(destPath + "\\WAdapterConfig.exe"))
                    {
                        if (!System.IO.File.Exists(destPath + "\\QoS\\WAdapterConfig.exe"))
                        {
                            CopyFileSafe(destPath + "\\WAdapterConfig.exe", destPath + "\\QoS\\WAdapterConfig.exe");
                        }
                    }

                    if (System.IO.File.Exists(destPath + "\\Config.xml"))
                    {
                        if (!System.IO.File.Exists(destPath + "\\QoS\\Config.xml"))
                        {
                            CopyFileSafe(destPath + "\\Config.xml", destPath + "\\QoS\\Config.xml");
                        }
                    }

                    updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\WLANInfoConfig\VACFlag.txt", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\VACFlag.txt");
                    if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                    updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\WLANInfoConfig\WLANInfo.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\WLANInfo.exe");
                    if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                    updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\WLANInfoConfig\WLANInfoConfig.xml", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\WLANInfoConfig.xml");
                    if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                    updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\QoS\Config.xml", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\QoS\\Config.xml");
                    if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                    updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\QoS\WAdapterConfig.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\QoS\\WAdapterConfig.exe");
                    if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                    updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\Tools\WiFiVwr.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\WiFiVwr.exe");
                    if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                    updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\Tools\DrvCtrl.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\ITTools\\DrvCtrl.exe");
                    if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                    updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\Tools\GetDeviceInfo.cmd", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\ITTools\\GetDeviceInfo.cmd");
                    if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                    updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\Tools\WLANLogs.cmd", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\WLANLogs.cmd");
                    if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());

                    if (bFullRun)
                    {
                        updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\Tools\UPDLocalFile.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\UPDLocalFile.exe");
                        if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                        //CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\Tools\GetWLANDriversInfo.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\GetWLANDriversInfo.exe");
                        //updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\Tools\WLANInfo.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\WLANInfo.exe");
                        //if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                        updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\Tools\GetConnLogs.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\GetConnLogs.exe");
                        if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                        //updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\QoS\AdapterInfo.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\QoS\\AdapterInfo.exe");
                        //if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                        //updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\Tools\WLANClient.chm", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\WLANClient.chm");
                        //if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                        updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\Tools\PNPDrvDeploy.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\ITTools\\PNPDrvDeploy.exe");
                        if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());
                        updateStatus = CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\Tools\DriverInfo.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\ITTools\\DriverInfo.exe");
                        if (bVerboseLog) Log("\t- update: " + updateStatus.ToString());

                        CreateWLANClientShortcuts();
                    }

                }

                GetCurrentDeploymentStatus();


                /// check QoS status
                /// run only every other 3rd day.. no need to run daily
                /// 
                if (bFullRun)
                {
                    if (System.IO.File.Exists(Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\QoS\\WAdapterConfig.exe"))
                    {
                        ThreadParameters checkQoS = new ThreadParameters(Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\QoS\\WAdapterConfig.exe", "/get", ProcessWindowStyle.Hidden, true);
                        long qosres = RunShellInWorkerThread(checkQoS);
                    }
                }


                /// next run NETS to get the driver version
                /// 
                RunNetSHToGetDriverVersion();
                //RunDevConToGetDevices();

                if (System.IO.File.Exists("C:\\temp\\netshdrivers.txt"))
                {
                    StreamReader sr = new System.IO.StreamReader("C:\\temp\\netshdrivers.txt");
                    string content = sr.ReadToEnd();
                    sr.Close();
                    ReadNETSHInfoFromFile(content);
                }


                //if (bFullRun)
                //    RunNetSHToGetWLANReport();
                ConfigureCustomEventLogView();
                SetRegkeyAdapterpath();

                CheckUpdateStatus();

                SaveWLANInfoToRegistry();

                if (!bSilentRun)
                    Console.WriteLine(inforesult);

                if (bWait)
                {
                    Console.WriteLine("Hit any key to close");
                    Console.ReadKey();
                }
                retCode = 0;
            }
            catch (Exception ex)
            {
                Log("Error in main module: " + ex.Message, true);
                retCode = 1;
            }

            PrepareUpdateAfterTask();
            return retCode;
        }

        public static void CreateWLANClientShortcuts()
        {
            try
            {
                string strWLANLocalPath = Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN";
                string strITToolsLocalPath = Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\ITTools";
                string linkFolder = string.Empty; //Environment.GetFolderPath(Environment.SpecialFolder.StartMenu) + @"\Programs\Intel\Intel WiDi PoC";
                RegistryKey shellFolders = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", RegistryKeyPermissionCheck.ReadSubTree);
                if (shellFolders == null)
                    linkFolder = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Intel\IT Client";
                else
                    linkFolder = shellFolders.GetValue("Common Programs", @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs").ToString() + @"\Intel\IT Client";

                try
                {
                    if (!Directory.Exists(linkFolder))
                        Directory.CreateDirectory(linkFolder);
                }
                catch (Exception pEx)
                {
                    Log("No access to create program folder in common programs space (" + pEx.Message + ")... try user space", true);
                    linkFolder = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu) + @"\Programs\Intel\IT Client";
                    if (!Directory.Exists(linkFolder))
                        Directory.CreateDirectory(linkFolder);
                }

                string shortCutFile = linkFolder + "\\WLAN Client User Guide.lnk";
                if (System.IO.File.Exists(strWLANLocalPath + @"\WLANClient.chm"))
                    MKShortcutFile(shortCutFile, strWLANLocalPath + @"\WLANClient.chm", string.Empty, "Intel WLAN Client - User Guide");

                shortCutFile = linkFolder + "\\Wireless Client Info.lnk";
                if (System.IO.File.Exists(strWLANLocalPath  + "\\WLANInfo.exe"))
                    MKShortcutFile(shortCutFile, strWLANLocalPath  + "\\WLANInfo.exe", string.Empty, "Intel IT Wireless Client Informaiton");

                shortCutFile = linkFolder + "\\Collect Logs For Wireless Troubleshooting.lnk";
                if (System.IO.File.Exists(strWLANLocalPath  + "\\GetConnLogs.exe"))
                    MKShortcutFile(shortCutFile, strWLANLocalPath  + "\\GetConnLogs.exe", string.Empty, "Collects System Data For Wireless Troubleshooting");

                shortCutFile = linkFolder + "\\Collect Logs For Device Troubleshooting.lnk";
                if (System.IO.File.Exists(strITToolsLocalPath + "\\GetDeviceInfo.cmd"))
                    MKShortcutFile(shortCutFile, strITToolsLocalPath + "\\GetDeviceInfo.cmd", string.Empty, "Collects System Data For Device Troubleshooting");

                shortCutFile = linkFolder + "\\Wireless Client QoS.lnk";
                if (System.IO.File.Exists(strWLANLocalPath + "\\QoS\\WAdapterConfig.exe"))
                    MKShortcutFile(shortCutFile, strWLANLocalPath + "\\QoS\\WAdapterConfig.exe", "/show /xml:C:\\ProgramData\\Intel\\WLAN\\QoS\\Config.xml", "Intel IT Wireless Client QoS Status");

                string OldLinkFolder = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Intel\WLAN Client";
                if (Directory.Exists(OldLinkFolder))
                    Directory.Delete(OldLinkFolder);

            }
            catch (Exception ex)
            {
                Log(ex.StackTrace);
            }
        }

        public static void MKShortcutFile(string shortcut, string target, string arguments, string description)
        {
            string linkFile = Path.GetTempPath() + "lLink.vbs";
            linkFile = GetNewFileForScript(linkFile, ".vbs");
            iLogFile lFile = new iLogFile(linkFile);
            lFile.LogInfo("On Error Resume Next");
            lFile.LogInfo("Dim objWSHShell, objSC");
            lFile.LogInfo("set objWSHShell = CreateObject(\"WScript.Shell\")");
            lFile.LogInfo("set objSC = objWSHShell.CreateShortcut(\"" + shortcut + "\")");
            lFile.LogInfo("objSC.Description = \"" + description + "\"");
            if (!arguments.Equals(string.Empty))
                lFile.LogInfo("objSC.Arguments = \"" + arguments + "\"");
            //lFile.LogInfo("objSC.IconLocation = \"" + System.Windows.Forms.Application.ExecutablePath + ",0\"");
            lFile.LogInfo("objSC.TargetPath = \"" + target + "\"");
            lFile.LogInfo("objSC.Save");
            
            RunShell("wscript.exe", "\"" + lFile.logFullName + "\"", ProcessWindowStyle.Hidden, true);
            File.Delete(linkFile);
        }

        public static string GetNewFileForScript(string fileName, string extension)
        {
            string functionReturnValue = fileName;
            if (File.Exists(fileName))
            {
                FileInfo fi = new FileInfo(fileName);
                int cnt = 0;
                while (cnt < 10)
                {
                    Thread.Sleep(400);
                    if (fi.LastAccessTime.AddMilliseconds(200) < DateTime.Now)
                    {
                        try
                        {
                            if (Program.bVerboseLog) Log(fileName + " no longer in use - delete now", true);
                            fi.Delete();
                            cnt = 10;
                        }
                        catch (System.Exception ex)
                        {
                            Log(ex.StackTrace);
                        }
                    }
                    else
                        if (Program.bVerboseLog) Log(fileName + " is still in use... wait before delete", true);
                    cnt += 1;
                }
            }
            if (File.Exists(fileName))
            {
                if (Program.bVerboseLog) Log(fileName + " is still in use... will create a temp file", true);
                functionReturnValue = System.IO.Path.GetTempFileName();
                functionReturnValue += extension;
            }
            return functionReturnValue;
        }

        public static long RunShell(string strCmd, string strArgs, ProcessWindowStyle WinStyle, bool Wait)
        {
            try
            {
                Process proc = null;
                ProcessStartInfo procInfo = new ProcessStartInfo(strCmd);
                procInfo.Arguments = strArgs;
                procInfo.WindowStyle = WinStyle;
                procInfo.ErrorDialog = true;
                procInfo.WorkingDirectory = Environment.GetEnvironmentVariable("Temp");

                proc = Process.Start(procInfo);

                if (Wait)
                {
                    while (!proc.HasExited)
                    {
                        Thread.Sleep(1000);
                    }
                    return proc.ExitCode;
                }
                else
                    return -1;

            }
            catch (System.Exception e)
            {
                Log("RunShell: " + e.StackTrace);
                return -1;
            }
        }





        public static void GetCurrentDeploymentStatus()
        {
            string file2read = Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\VACFlag.txt";
            if (File.Exists(file2read))
            {
                string fileContent = File.ReadAllText(file2read);
                using (System.IO.StringReader rdr = new StringReader(fileContent))
                {
                    string line;
                    bVACGreenLight = false;
                    while ((line = rdr.ReadLine()) != null)
                    {
                        line = line.Trim();
                        if (line.ToLower().StartsWith("check"))
                            bVACGreenLight = true;
                        if (line.ToLower().StartsWith("run"))
                            bVACGreenLight = true;
                    }
                }

            }
            //using (WebClient wc = new WebClient())
            //{
            //    //wc.DownloadProgressChanged += wc_DownloadProgressChanged;
            //    wc.DownloadFile(
            //        // Param1 = Link of file
            //        new System.Uri("https://intel.service-now.com/kb_view.do?sysparm_article=KB000311502"),
            //        // Param2 = Path to save
            //        "C:\\temp\\WLANCurrent.txt"
            //    );
            //}
        }


        public static void CheckUpdateStatus()
        {
            try
            {
                if (System.IO.File.Exists(Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\WLANInfoConfig.xml"))
                {
                    /// run the wlaninfo routines to check if there is an update to the client
                    /// 
                    wirelessUpdateStatus = CheckIfWirelessCanBeUpdated();
                    regKeyPackage = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
                    if (regKeyPackage == null)
                    {
                        Registry.LocalMachine.CreateSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
                        regKeyPackage = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
                    }

                    if (wlaninfo == null)
                    {
                        wlaninfo = new WirelessLANInfo();
                    }

                    if (wirelessUpdateStatus == WLANClientUpdateStatus.NOT_APPLICABLE)
                    {
                        Log("Wireless LAN Client is not applicable", true);
                        regKeyPackage.SetValue("PassThruPath", "not applicable");
                        regKeyPackage.SetValue("UpdateCMD", "not applicable");
                        regKeyPackage.SetValue("UpdateCMDArgs", "not applicable");
                        regKeyPackage.SetValue("Current Available Version", "not applicable");
                        regKeyPackage.SetValue("isCurrent", "not applicable");
                        regKeyPackage.SetValue("VACRun", 0, RegistryValueKind.DWord);
                    }
                    else
                    {
                        if (wirelessUpdateStatus == WLANClientUpdateStatus.UPDATE_AVAILABLE)
                            Log("Update available: " + wlaninfo.PassThruPath, true);
                        else
                            Log("Wireless LAN Client version found: " + wlaninfo.AppVersionFound, true);
                        regKeyPackage.SetValue("PassThruPath", ResolveVariableInString(wlaninfo.PassThruPath));
                        regKeyPackage.SetValue("UpdateCMD", ResolveVariableInString(wlaninfo.PathToInstaller) + "\\" + wlaninfo.Installer);
                        regKeyPackage.SetValue("UpdateCMDArgs", ResolveVariableInString(wlaninfo.CmdArguments));
                        regKeyPackage.SetValue("Current Available Version", wlaninfo.AppVersionCurrentRelease);
                        regKeyPackage.SetValue("isCurrent", wirelessUpdateStatus.ToString());
                        if (bVACGreenLight)
                            regKeyPackage.SetValue("VACRun", 1, RegistryValueKind.DWord);
                        else
                            regKeyPackage.SetValue("VACRun", 0, RegistryValueKind.DWord);
                    }

                }

            }
            catch (Exception ex)
            {
                Log("CheckUpdateStatus error: " + ex.Message, true);
            }
        }

        public static void ConfigureTheScheduledTask()
        {
            string schtasksSCMD = Environment.GetFolderPath(Environment.SpecialFolder.System) + "\\SCHTASKS.exe";
            if (System.IO.File.Exists(schtasksSCMD))
            {
                try
                {
                    string taskFile = Environment.GetFolderPath(Environment.SpecialFolder.System) + "\\Tasks\\Intel\\WLAN Client Info";
                    if (System.IO.File.Exists(taskFile))
                    {
                        CopyFileSafe(taskFile, "C:\\temp\\tsk");
                        StreamReader sr = new StreamReader("C:\\temp\\tsk");
                        string content = sr.ReadToEnd();
                        if (content.ToLower().Contains("drivers\\wlan"))
                            bFroceTaskCreation = true;
                        //if (content.ToUpper().Contains("NT AUTHORITY\\SYSTEM"))
                        //    bFroceTaskCreation = true;
                    }
                    else
                        bFroceTaskCreation = true;
                    //ITaskService tsk = new TaskScheduler.TaskScheduler();
                    //tsk.Connect();
                    //ITaskFolder tskFolder = tsk.GetFolder("Intel");
                    //foreach (IRegisteredTask task in tskFolder.GetTasks(1))
                    //{
                    //    Debug.Print(task.Name);
                    //    if (task.Name.ToLower().Contains("wlan client info"))
                    //    {
                    //        foreach (IAction action in task.Definition.Actions)
                    //        {
                    //            if (action.Id.Contains("ProgramData")) Debug.Print("Found it");
                    //            if (action.ToString().ToLower().Contains(@"programdata\intel\wlan"))
                    //            //if (task.Path.ToLower().Contains(@"programdata\intel\wlan"))
                    //            {
                    //                Debug.Print("task created and in the right place");
                    //            }
                    //            else
                    //            {
                    //                Debug.Print("task created but not in the right place");
                    //                bFroceTaskCreation = true;
                    //            }
                    //        }
                    //    }
                    //}


                    ThreadParameters createTask = new ThreadParameters();
                    string schtasksArgs = string.Empty;
                    bool bCreateTask = false;
                    schtasksArgs = " /Query /TN \"Intel\\WLAN Client Info\"";
                    createTask = new ThreadParameters(schtasksSCMD, schtasksArgs, ProcessWindowStyle.Hidden, true);
                    long result = RunShellInWorkerThread(createTask);
                    if (result == 0) // the task does exist
                    {
                        if (bFroceTaskCreation) // delete it first if already in place... 
                        {
                            schtasksArgs = " /Delete /TN \"Intel\\WLAN Client Info\" /F";
                            createTask = new ThreadParameters(schtasksSCMD, schtasksArgs, ProcessWindowStyle.Hidden, true);
                            result = RunShellInWorkerThread(createTask);
                            bCreateTask = true;
                            Thread.Sleep(1500);
                        }
                        else
                            bCreateTask = false;
                    }
                    else
                        bCreateTask = true;

                    if (bCreateTask)
                    {
                        if (CreateTaskXMLFile())
                        {
                            schtasksArgs = " /Create /TN \"Intel\\WLAN Client Info\" /XML \"" + Environment.GetEnvironmentVariable("SystemDrive") + "\\Temp\\WLAN Client Info.xml\" /RU SYSTEM /F";
                            Thread.Sleep(500);
                            //if (System.IO.File.Exists(Environment.GetEnvironmentVariable("SystemDrive") + "\\Temp\\WLAN Client Info.xml"))
                            //    System.IO.File.Delete(Environment.GetEnvironmentVariable("SystemDrive") + "\\Temp\\WLAN Client Info.xml");
                        }
                        else
                        {
                            schtasksArgs = " /Create /SC DAILY /TN \"Intel\\WLAN Client Info\" /TR \"" + Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\GetWLANDriversInfo.exe /s\" /ST 02:00 /RU SYSTEM /F";
                        }
                        createTask = new ThreadParameters(schtasksSCMD, schtasksArgs, ProcessWindowStyle.Hidden, true);
                        result = RunShellInWorkerThread(createTask);
                    }
                }
                catch { }
            }

        }

        public static RunningPlatform SizeOfIntPtr()
        {
            RunningPlatform functionReturnValue;
            switch (IntPtr.Size)
            {
                case 4:
                    functionReturnValue = RunningPlatform._32;
                    break;
                case 8:
                    functionReturnValue = RunningPlatform._64;
                    break;
                default:
                    functionReturnValue = RunningPlatform.unknown;
                    break;
            }
            return functionReturnValue;
        }

        public static bool CreateTaskXMLFile()
        {
            bool functionReturnValue = false;

            try
            {
                string cmd2Run = Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\GetWLANDriversInfo.exe";
                StringBuilder sb = new StringBuilder();

                DateTime tomorrow = DateTime.Now.AddDays(1);

                sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-16\"?>");
                sb.AppendLine("<Task version=\"1.2\" xmlns=\"http://schemas.microsoft.com/windows/2004/02/mit/task\">");
                sb.AppendLine("  <RegistrationInfo>");
                sb.AppendLine("    <Date>" + DateTime.Now.Year + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0') + "T" + DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Second.ToString().PadLeft(2, '0') + "</Date>");
                sb.AppendLine("    <Author>" + Environment.UserName + "</Author>");
                sb.AppendLine("  </RegistrationInfo>");
                sb.AppendLine("  <Triggers>");
                sb.AppendLine("    <CalendarTrigger>");
                sb.AppendLine("      <StartBoundary>" + tomorrow.Year + "-" + tomorrow.Month.ToString().PadLeft(2, '0') + "-" + tomorrow.Day.ToString().PadLeft(2, '0') + "T" + tomorrow.Hour.ToString().PadLeft(2, '0') + ":" + tomorrow.Minute.ToString().PadLeft(2, '0') + ":" + tomorrow.Second.ToString().PadLeft(2, '0') + "</StartBoundary>");
                sb.AppendLine("      <Enabled>true</Enabled>");
                sb.AppendLine("      <ScheduleByDay>");
                sb.AppendLine("        <DaysInterval>1</DaysInterval>");
                sb.AppendLine("      </ScheduleByDay>");
                sb.AppendLine("    </CalendarTrigger>");
                sb.AppendLine("  </Triggers>");
                sb.AppendLine("  <Principals>");
                sb.AppendLine("    <Principal id=\"Author\">");
                if(CurrentLogedOnUser.ToLower().Contains("amr\\") || CurrentLogedOnUser.ToLower().Contains("gar\\") || CurrentLogedOnUser.ToLower().Contains("ger\\") || CurrentLogedOnUser.ToLower().Contains("ccr\\"))
                {
                    sb.AppendLine("      <UserId>" + CurrentLogedOnUser + "</UserId>");
                }
                //sb.AppendLine("      <UserId>S-1-5-20</UserId>");
                //sb.AppendLine("      <UserId></UserId>");
                sb.AppendLine("      <RunLevel>HighestAvailable</RunLevel>");
                sb.AppendLine("    </Principal>");
                sb.AppendLine("  </Principals>");
                sb.AppendLine("  <Settings>");
                sb.AppendLine("    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>");
                sb.AppendLine("    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>");
                sb.AppendLine("    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>");
                sb.AppendLine("    <AllowHardTerminate>true</AllowHardTerminate>");
                sb.AppendLine("    <StartWhenAvailable>true</StartWhenAvailable>");
                sb.AppendLine("    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>");
                sb.AppendLine("    <IdleSettings>");
                sb.AppendLine("      <StopOnIdleEnd>true</StopOnIdleEnd>");
                sb.AppendLine("      <RestartOnIdle>false</RestartOnIdle>");
                sb.AppendLine("    </IdleSettings>");
                sb.AppendLine("    <AllowStartOnDemand>true</AllowStartOnDemand>");
                sb.AppendLine("    <Enabled>true</Enabled>");
                sb.AppendLine("    <Hidden>false</Hidden>");
                sb.AppendLine("    <RunOnlyIfIdle>false</RunOnlyIfIdle>");
                sb.AppendLine("    <WakeToRun>false</WakeToRun>");
                sb.AppendLine("    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>");
                sb.AppendLine("    <Priority>7</Priority>");
                sb.AppendLine("  </Settings>");
                sb.AppendLine("  <Actions Context=\"Author\">");
                sb.AppendLine("    <Exec>");
                sb.AppendLine("      <Command>" + cmd2Run + "</Command>");
                sb.AppendLine("      <Arguments>/s</Arguments>");
                sb.AppendLine("    </Exec>");
                sb.AppendLine("  </Actions>");
                sb.AppendLine("</Task>");


                if (!System.IO.Directory.Exists(Environment.GetEnvironmentVariable("SystemDrive") + "\\Temp"))
                    System.IO.Directory.CreateDirectory(Environment.GetEnvironmentVariable("SystemDrive") + "\\Temp");
                if (System.IO.File.Exists(Environment.GetEnvironmentVariable("SystemDrive") + "\\Temp\\WLAN Client Info.xml"))
                    System.IO.File.Delete(Environment.GetEnvironmentVariable("SystemDrive") + "\\Temp\\WLAN Client Info.xml");

                using (StreamWriter outfile = new StreamWriter(Environment.GetEnvironmentVariable("SystemDrive") + "\\Temp\\WLAN Client Info.xml"))
                {
                    outfile.Write(sb.ToString());
                }
                functionReturnValue = true;
            }
            catch
            {
                functionReturnValue = false;
            }

            if (System.IO.File.Exists(Environment.GetEnvironmentVariable("SystemDrive") + "\\Temp\\WLAN Client Info.xml"))
                functionReturnValue = true;
            
            return functionReturnValue;
        }

        public static void SetRegkeyAdapterpath()
        {
            try
            {
                AllWirelessConnections = GetAllWirelessConnections();
                if (AllWirelessConnections.Count < 1)
                {
                    if (System.IO.File.Exists("C:\\ProgramData\\Intel\\WLAN\\devcon.exe"))
                    {
                        if (bVerboseLog) 
                            Console.WriteLine("No wireless interfaces detected... force a hardware rescan via devcon", true);
                        try
                        {
                            ThreadParameters devConThread = new ThreadParameters();
                            devConThread.CMDLine = "C:\\ProgramData\\Intel\\WLAN\\devcon.exe";
                            devConThread.ProcWindowStyle = ProcessWindowStyle.Hidden;
                            devConThread.WaitForThread = true;
                            devConThread.CMDLineArguments = "rescan";
                            RunShellInWorkerThread(devConThread);
                            Thread.Sleep(2000);
                            AllWirelessConnections = GetAllWirelessConnections();
                        }
                        catch (Exception ex)
                        {
                            if (bVerboseLog) Console.WriteLine("Failure in rescannign the hardware: " + ex.Message, true);
                        }
                    }
                }

                if (AllWirelessConnections.Count > 0)
                {
                    foreach (KeyValuePair<string, string> entry in AllWirelessConnections)
                    {
                        //RegistryKey RegKeyAdapter = GetAdapterRegistryKey(entry);
                        if (bVerboseLog) Log("Checking reg space for: " + entry.Key + "|" + entry.Value, true);
                        // RegKeyAdapter = GetAdapterRegistryKey(entry); // old style detection 
                        RegkeyAdapterPath = GetWirelessAdapterRegistryKey(entry).ToString(); // new style detection - using GUID / NetCfgInstanceID
                    }
                }
                else
                {
                    if (bVerboseLog) Console.WriteLine("No interfaces detected post install... will require a reboot before running the post install configurations", true);
                }
            }
            catch { }
        }

        #region New Style registry key detection

        public static RegistryKey GetWirelessAdapterRegistryKey(KeyValuePair<string, string> netcon)
        {
            RegistryKey functionReturnValue = null;
            if (netcon.Value.ToLower().Contains("virtual"))
            {
                if (bVerboseLog) Log("Virtual Interface " + netcon.Value + " - skip searching for PnPInstaceID", true);
            }
            else
            {
                string[] deviceIDs = netcon.Value.Split('|');
                if (deviceIDs.Count() == 2)
                {
                    functionReturnValue = GetWirelessAdapterRegistryKey(deviceIDs[0], deviceIDs[1]);
                }
            }

            return functionReturnValue;
        }

        public static RegistryKey GetWirelessAdapterRegistryKey(string adapter, string InstanceGUID)
        {
            RegistryKey functionReturnValue = null;
            if (adapter.Contains("#"))
                adapter = adapter.Substring(0, adapter.IndexOf("#")).Trim();
            wirelessAdapter = adapter;
            try
            {
                Log("Attempt to read wireless adapter registry settings for " + adapter + " (GUID: " + InstanceGUID + ")", true);
                bool bFoundAMatch = false;
                RegistryKey regKeyClass = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}", RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
                if (regKeyClass == null)
                {
                    Log("Cannot read adapters registry space to apply the registry fix... access denied", true);
                }
                else
                {
                    string[] strNetAdapters = regKeyClass.GetSubKeyNames();
                    foreach (string strAdapterID in strNetAdapters)
                    {
                        if (strAdapterID.ToLower().Equals("properties"))
                            continue;
                        RegistryKey keyAdapter = regKeyClass.OpenSubKey(strAdapterID, RegistryKeyPermissionCheck.ReadWriteSubTree);
                        if (keyAdapter == null)
                        {
                            Log("Cannot read adapter registry space to apply the registry configuration... access dennied", true);
                        }
                        else
                        {
                            string strDescription = keyAdapter.GetValue("DriverDesc", string.Empty).ToString();
                            if (strDescription.Contains("#"))
                                strDescription = strDescription.Substring(0, adapter.IndexOf("#")).Trim();
                            string descCompare = strDescription.ToLower().Replace(" ", "").Replace("\"", "");
                            string adapterCompare = adapter.ToLower().Replace(" ", "").Replace("\"", "");
                            if (bVerboseLog) Log("Compare device ID in (" + strAdapterID + "):\r\n\t\t" + descCompare + "\r\n\t\t" + adapterCompare);
                            if (descCompare.Equals(adapterCompare))
                            {
                                string GUID_ID = keyAdapter.GetValue("NetCfgInstanceId", string.Empty).ToString();
                                if (!GUID_ID.Equals(string.Empty))
                                {
                                    string deviceIDComp = InstanceGUID.ToLower().Replace(" ", "").Replace("\"", "");
                                    string instaceIDComp = GUID_ID.ToLower().Replace(" ", "").Replace("\"", "");
                                    if (bVerboseLog) Log("\r\n\t\t" + deviceIDComp + "\r\n\t\t" + instaceIDComp);
                                    if (deviceIDComp.Equals(instaceIDComp))
                                    {
                                        bFoundAMatch = true;
                                        Log("Found a match in " + keyAdapter.ToString());
                                        functionReturnValue = keyAdapter;
                                        /// now get the DeviceinstanceID and see if it's an Intel adapter... 
                                        /// 
                                        DeviceInstanceID = keyAdapter.GetValue("DeviceInstanceID", string.Empty).ToString();
                                        if (DeviceInstanceID.ToUpper().Contains("VEN_8086"))
                                            bIsIntelAdapter = true;
                                        else
                                        {
                                            Log("Device Instance ID (" + DeviceInstanceID + ") not matching an Intel adapter");
                                            bIsIntelAdapter = false;
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    if (!bFoundAMatch)
                    {
                        Log("No wireless adapter detected", true);
                    }
                }
            }
            catch (Exception regEx)
            {
                Log("Cannot read adapters registry space: " + regEx.Message, true);

            }
            return functionReturnValue;
        }


        #endregion

        public static Dictionary<string, string> GetAllWirelessConnections()
        {
            Dictionary<string, string> awc = new Dictionary<string, string>();

            try
            {
                //Display current network adapter states
                NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
                string macAddress = string.Empty;
                foreach (NetworkInterface n in adapters)
                {
                    macAddress = n.GetPhysicalAddress().ToString().Replace(":", "");
                    //Log(string.Format("\tSpeed: {0}", n.Speed));
                    if (n.NetworkInterfaceType == NetworkInterfaceType.Wireless80211)
                    {
                        if (n.Description.ToLower().Contains("virtual") || n.Description.ToLower().Contains("bluetooth"))
                        {
                            if (bVerboseLog) Log("Virtual Adapter - not adding " + n.Name + " to wireless collection", true);
                        }
                        else
                        {
                            if (bVerboseLog) Log("Adding " + n.Name + " to wireless collection", true);
                            awc.Add(n.Name, n.Description + "|" + n.Id);
                            interfaceName = n.Name;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while reading wireless conenctions: " + ex.Message);
            }
            return awc;
        }

        public static RegistryKey GetAdapterRegistryKey(KeyValuePair<string, string> netcon)
        {
            RegistryKey functionReturnValue = null;
            string PnPInstanceID = GetPnPInstanceID(netcon.Key);
            if (!PnPInstanceID.Equals(string.Empty))
            {
                functionReturnValue = AdapterRegistryKey(netcon.Value, PnPInstanceID);
                RegkeyAdapterPath = functionReturnValue.ToString();
            }
            return functionReturnValue;
        }

        public static RegistryKey AdapterRegistryKey(string adapter, string deviceID)
        {
            RegistryKey functionReturnValue = null;
            if (adapter.Contains("#"))
                adapter = adapter.Substring(0, adapter.IndexOf("#")).Trim();
            try
            {
                bool bFoundAMatch = false;
                RegistryKey regKeyClass = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}", RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
                if (regKeyClass == null)
                {
                    Console.WriteLine("Cannot read adapters registry space to apply the registry fix... access denied", true);
                }
                else
                {
                    string[] strNetAdapters = regKeyClass.GetSubKeyNames();
                    foreach (string strAdapterID in strNetAdapters)
                    {
                        if (strAdapterID.ToLower().Equals("properties"))
                            continue;
                        RegistryKey keyAdapter = regKeyClass.OpenSubKey(strAdapterID, RegistryKeyPermissionCheck.ReadWriteSubTree);
                        if (keyAdapter == null)
                        {
                            if (bVerboseLog) Console.WriteLine("Cannot read adapter registry space to apply the registry configuration... access dennied");
                        }
                        else
                        {
                            string strDescription = keyAdapter.GetValue("DriverDesc", string.Empty).ToString();
                            if (strDescription.Contains("#"))
                                strDescription = strDescription.Substring(0, adapter.IndexOf("#")).Trim();
                            string descCompare = strDescription.ToLower().Replace(" ", "").Replace("\"", "");
                            string adapterCompare = adapter.ToLower().Replace(" ", "").Replace("\"", "");
                            if (bVerboseLog) Console.WriteLine("Compare device ID in (" + strAdapterID + "):\r\n\t\t" + descCompare + "\r\n\t\t" + adapterCompare);
                            if (descCompare.Equals(adapterCompare))
                            {
                                string deviceInstanceID = keyAdapter.GetValue("DeviceInstanceID", string.Empty).ToString();
                                if (!deviceInstanceID.Equals(string.Empty))
                                {
                                    string deviceIDComp = deviceID.ToLower().Replace(" ", "").Replace("\"", "");
                                    string instaceIDComp = deviceInstanceID.ToLower().Replace(" ", "").Replace("\"", "");
                                    if (bVerboseLog) Console.WriteLine("\r\n\t\t" + deviceIDComp + "\r\n\t\t" + instaceIDComp);
                                    if (deviceIDComp.Equals(instaceIDComp))
                                    {
                                        bFoundAMatch = true;
                                        if (bVerboseLog) Console.WriteLine("Found a match in " + keyAdapter.ToString());
                                        functionReturnValue = keyAdapter;
                                    }
                                }
                            }
                        }
                    }
                    if (!bFoundAMatch)
                    {
                        if (bVerboseLog) Console.WriteLine("No wireless adapter detected", true);
                    }
                }
            }
            catch (Exception regEx)
            {
                if (bVerboseLog) Console.WriteLine("Cannot read adapters registry space: " + regEx.Message);
            }
            return functionReturnValue;
        }

        public static string GetPnPInstanceID(string ConnectionName)
        {
            string functionReturnValue = string.Empty;
            try
            {
                /// find the connection name in 
                /// HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}
                /// 
                RegistryKey regNet = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}", RegistryKeyPermissionCheck.ReadSubTree);
                if (regNet == null)
                {
                    if (bVerboseLog) Console.WriteLine(@"Cannot read network registry info in SYSTEM\CurrentControlSet\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}");
                }
                else
                {
                    string[] netSubKeys = regNet.GetSubKeyNames();
                    foreach (string netsubKey in netSubKeys)
                    {
                        RegistryKey netConnection = regNet.OpenSubKey(netsubKey + "\\Connection", RegistryKeyPermissionCheck.ReadSubTree);
                        if (netConnection != null)
                        {
                            string connName = netConnection.GetValue("Name", string.Empty).ToString();
                            if (connName.ToLower().Equals(ConnectionName.ToLower()))
                            {
                                string pnpInstanceID = netConnection.GetValue("PnpInstanceID", string.Empty).ToString();
                                if (pnpInstanceID.ToLower().Contains("pci\\ven"))
                                {
                                    if (!pnpInstanceID.Equals(string.Empty))
                                    {
                                        functionReturnValue = pnpInstanceID;
                                        break;
                                    }
                                }
                            }
                        }
                        else
                            Debug.Print(@"Skipping ...\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}\" + netsubKey + " - No connection reg key");
                    }
                }
            }
            catch (Exception ex)
            {
                if (bVerboseLog) Console.WriteLine("Error while getting PnP Instance ID: " + ex.Message);
            }

            return functionReturnValue;
        }

        public static void SaveWLANInfoToRegistry()
        {

            inforesult = string.Empty;
            try
            {
                /// interfaceName
                /// wirelessAdapter
                /// wirelessDriverVersion
                /// wirelessDriverFile
                /// 

                //if (interfaceName.Equals(string.Empty))
                //    interfaceName = "cannot determine";

                //if (wirelessAdapter.Equals(string.Empty))
                //    wirelessAdapter = "cannot determine";

                //if (wirelessDriverVersion.Equals(string.Empty))
                //    wirelessDriverVersion = "cannot determine";

                //if (wirelessDriverFile.Equals(string.Empty))
                //    wirelessDriverFile = "cannot determine";


                RegistryKey regKeyPackage = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (regKeyPackage == null)
                {
                    Registry.LocalMachine.CreateSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
                    regKeyPackage = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
                }
                if (!interfaceName.Equals(string.Empty))
                {
                    regKeyPackage.SetValue("InterfaceName", interfaceName, RegistryValueKind.String);
                    Log("Interface Name             : " + interfaceName);
                    inforesult += "\r\nInterface Name             : " + interfaceName;
                    //if (!bSilentRun)
                    //    Console.WriteLine("Interface Name           : " + interfaceName);
                }
                if (!wirelessAdapter.Equals(string.Empty))
                {
                    regKeyPackage.SetValue("WirelessAdapter", wirelessAdapter, RegistryValueKind.String);
                    Log("Wireless Adapter           : " + wirelessAdapter);
                    inforesult += "\r\nWireless Adapter           : " + wirelessAdapter;
                    //if (!bSilentRun)
                    //    Console.WriteLine("Wireless Adapter         : " + wirelessAdapter);
                }
                if (!wirelessDriverFile.Equals(string.Empty))
                {
                    regKeyPackage.SetValue("WirelessDriverFile", wirelessDriverFile, RegistryValueKind.String);
                    Log("Wireless Driver            : " + wirelessDriverFile);
                    inforesult += "\r\nWireless Driver            : " + wirelessDriverFile;
                    //if (!bSilentRun)
                    //    Console.WriteLine("Wireless Driver          : " + wirelessDriverFile);
                }
                if (!wirelessDriverVersion.Equals(string.Empty))
                {
                    regKeyPackage.SetValue("WirelessDriverVersion", wirelessDriverVersion, RegistryValueKind.String);
                    Log("Wireless Driver Version    : " + wirelessDriverVersion);
                    inforesult += "\r\nWireless Driver Version    : " + wirelessDriverVersion;
                    //if (!bSilentRun)
                    //    Console.WriteLine("Wireless Driver Version  : " + wirelessDriverVersion);
                }
                if (!DeviceInstanceID.Equals(string.Empty))
                {
                    regKeyPackage.SetValue("DeviceInstanceID", DeviceInstanceID, RegistryValueKind.String);
                    Log("Device Instance ID         : " + DeviceInstanceID);
                    inforesult += "\r\nDevice Instance ID         : " + DeviceInstanceID;
                    //if (!bSilentRun)
                    //    Console.WriteLine("Wireless Driver Version  : " + wirelessDriverVersion);
                }
                if (!RegkeyAdapterPath.Equals(string.Empty))
                {
                    regKeyPackage.SetValue("RegkeyAdapterPath", RegkeyAdapterPath.Substring(RegkeyAdapterPath.LastIndexOf("\\") + 1), RegistryValueKind.String);
                }

                string PSVer = GetPROSetVersion();
                if (CurrentPROSetVersion.Equals(string.Empty))
                    regKeyPackage.SetValue("Current PROSet Version", "not installed", RegistryValueKind.String);
                else
                    regKeyPackage.SetValue("Current PROSet Version", CurrentPROSetVersion, RegistryValueKind.String);

                //if (!CurrentDriverPackage.Equals(string.Empty))
                ////    regKeyPackage.SetValue("Current Driver Package", "not installed", RegistryValueKind.String);
                ////else
                //    regKeyPackage.SetValue("Current Driver Package", CurrentDriverPackage, RegistryValueKind.String);

                RegistryKey regQoS = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Intel\Payloads\Wireless Adapter Configuration", RegistryKeyPermissionCheck.ReadSubTree);
                if (regQoS != null)
                {
                    string qosStatus = regQoS.GetValue("Status", "Not configured").ToString();
                    string qosVer = regQoS.GetValue("QoSVersion", "not found").ToString();
                    string qosDate = regQoS.GetValue("Last run DT", "not set").ToString();
                    if (qosVer.Equals("not found"))
                    {
                        Log("Wireless QoS Status      : " + qosStatus);
                        inforesult += "\r\nWireless QoS Status        : " + qosStatus;
                    }
                    else
                    {
                        if (qosDate.Equals("not set"))
                        {
                            Log("Wireless QoS Status        : " + qosStatus + " (v." + qosVer + ")");
                            inforesult += "\r\nWireless QoS Status        : " + qosStatus + " (v." + qosVer + ")";
                        }
                        else
                        {
                            Log("Wireless QoS Status        : " + qosStatus + " (v." + qosVer + " as of " + qosDate + ")");
                            inforesult += "\r\nWireless QoS Status        : " + qosStatus + " (v." + qosVer + " as of " + qosDate + ")";
                        }
                    }

                    Log("Wireless LAN Client Status : " + wirelessUpdateStatus.ToString());
                    inforesult += "\r\nWireless LAN Client Status : " + wirelessUpdateStatus.ToString();
                }

                /// for systems w/out wireless, this will never be written otherwise .. do it here.. 
                if (wirelessUpdateStatus == WLANClientUpdateStatus.NOT_APPLICABLE)
                {
                    regKeyPackage.SetValue("Status", "NOT_APPLICABLE", RegistryValueKind.String);
                    regKeyPackage.SetValue("isCurrent", "NOT_APPLICABLE", RegistryValueKind.String);
                }
                else
                {
                    regKeyPackage.SetValue("Status", wirelessUpdateStatus.ToString(), RegistryValueKind.String);
                }

                regKeyPackage.SetValue("GWInfoVer", System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(), RegistryValueKind.String);
                DateTime runtime = DateTime.Now;
                //DateTime jan180 = new DateTime(2000, 1, 1, 0, 0, 0);
                //DateTime jan180a = new DateTime(1980, 1, 1, 1, 0, 0);
                // Thu, 31 Dec 2015

                string runtimestring = runtime.ToUniversalTime().ToString("R");
                
                //runtime.Year.ToString() + "/" + runtime.Month.ToString() + "/" + runtime.Day.ToString();
                regKeyPackage.SetValue("GWLastRun", runtimestring, RegistryValueKind.String);
                //regKeyPackage.SetValue("test", jan180.ToString("R"), RegistryValueKind.String);
                //regKeyPackage.SetValue("testa", jan180a.ToString("R"), RegistryValueKind.String);
            }
            catch (Exception ex)
            {
                Log("Err in saving info to registry: " + ex.Message, true);
                RegistryKey regKeyPackage = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
                if (regKeyPackage == null)
                {
                    Registry.LocalMachine.CreateSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
                    regKeyPackage = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
                }
                regKeyPackage.SetValue("GWInfoVer", System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(), RegistryValueKind.String);
            }
        }

        public static void ReadNETSHInfoFromFile(string output)
        {

            using (System.IO.StringReader reader = new StringReader(output))
            {
                string line;
                bool bFilesLine = false;
                //Console.WriteLine(output);
                while ((line = reader.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (line.StartsWith("Interface name"))
                    {
                        interfaceName = line.Substring(line.IndexOf(":") + 1).Trim();
                    }
                    if (line.StartsWith("Driver"))
                    {
                        wirelessAdapter = line.Substring(line.IndexOf(":") + 1).Trim();
                    }
                    if (line.StartsWith("Version"))
                    {
                        wirelessDriverVersion = line.Substring(line.IndexOf(":") + 1).Trim();
                    }
                    if (bFilesLine)
                    {
                        wirelessDriverFile = line.Trim();
                        bFilesLine = false;
                    }
                    if (line.StartsWith("Files"))
                    {
                        bFilesLine = true;
                    }
                    //Console.WriteLine(line); 
                }
            }
        }

        public static string GetPROSetVersion()
        {
            string functionReturnValue = string.Empty;
            try
            {
                //AllInstalledApps.Sort();
                bool bPROSetInstalled = false;
                foreach (string app in AllInstalledApps)
                {
                    if (bVerboseLog) 
                    {
                        if (!bSilentRun)
                            Console.WriteLine("Checking installed app " + app);
                    }

                    string[] appInfo = app.Split('|');
                    if (appInfo[0].ToLower().EndsWith("proset/wireless software"))
                    {
                        bPROSetInstalled = true;
                        functionReturnValue = appInfo[1];
                        CurrentPROSetVersion = functionReturnValue;
                        break;
                    }
                    if (appInfo[0].ToLower().EndsWith("pro/wireless driver"))
                    {
                        CurrentDriverPackage = appInfo[1];
                        break;
                    }

                }
                if (!bPROSetInstalled)
                {
                    if (!CurrentDriverPackage.Equals(string.Empty))
                        functionReturnValue = CurrentDriverPackage;
                }
            }
            catch (Exception ex)
            {
                if (bVerboseLog)
                {
                    if (!bSilentRun)
                        Console.WriteLine(ex.Message);
                }
            }
            return functionReturnValue;
        }

        public static List<string> GetInstalledApplications()
        {
            string registryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            List<string> myApps = new List<string>();
            using (Microsoft.Win32.RegistryKey key = Registry.LocalMachine.OpenSubKey(registryKey))
            {
                (from a in key.GetSubKeyNames()
                 let r = key.OpenSubKey(a)
                 select new
                 {
                     UninstallEntry = r.GetValue("DisplayName") + "|" + r.GetValue("DisplayVersion") + "|" + a
                 }).ToList()
                    //.FindAll(c => c.UninstallEntry != null)
                    .FindAll(c => !c.UninstallEntry.ToString().StartsWith("|"))
                    .ForEach(c => myApps.Add(c.UninstallEntry.ToString()));
            }

            if (Program.platform == RunningPlatform._64)
            {
                registryKey = @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
                using (Microsoft.Win32.RegistryKey key = Registry.LocalMachine.OpenSubKey(registryKey))
                {
                    (from a in key.GetSubKeyNames()
                     let r = key.OpenSubKey(a)
                     select new
                     {
                         UninstallEntry = r.GetValue("DisplayName") + "|" + r.GetValue("DisplayVersion") + "|" + a
                     }).ToList()
                        //.FindAll(c => c.UninstallEntry != null)
                        .FindAll(c => !c.UninstallEntry.ToString().StartsWith("|"))
                        .ForEach(c => myApps.Add(c.UninstallEntry.ToString()));
                    //.ForEach(c => Debug.Print(c.UninstallEntry.ToString()));
                }
            }
            //sanityze myApps - remove entries that have no version
            IEnumerable<string> appsQuery =
                from app in myApps
                where !app.Contains("||")
                select app;
            List<string> functionReturnValue = new List<string>();
            foreach (string ap in appsQuery)
            {
                if (Program.bVerboseLog) Console.WriteLine("- GetInstalledApps: " + ap);
                functionReturnValue.Add(ap);
            }

            return functionReturnValue; ;
        }

        public static long RunShellInWorkerThread(ThreadParameters runArguments) // string strCmd, string strArgs, ProcessWindowStyle WinStyle, bool Wait)
        {
            Program.runShellArguments = new ThreadParameters(runArguments.CMDLine, runArguments.CMDLineArguments, runArguments.ProcWindowStyle, runArguments.WaitForThread);
            Thread workerThread = new Thread(new ThreadStart(RunShellAsWorkerThread));
            workerThread.Name = "Autorun as worker thread";
            workerThread.Start();
            while (workerThread.IsAlive)
            {
                Thread.Sleep(100);
            }

            if (bVerboseLog) Console.WriteLine("Result:" + Program.runShellArguments.ReturnCode.ToString());
            return Program.runShellArguments.ReturnCode;
        }

        private static void RunShellAsWorkerThread()
        {
            try
            {
                Process proc = null;
                ProcessStartInfo procInfo = new ProcessStartInfo(Program.runShellArguments.CMDLine);
                procInfo.Arguments = Program.runShellArguments.CMDLineArguments;
                procInfo.WindowStyle = Program.runShellArguments.ProcWindowStyle;
                procInfo.ErrorDialog = true;
                procInfo.UseShellExecute = false;
                procInfo.WorkingDirectory = Environment.GetEnvironmentVariable("Temp");
                procInfo.RedirectStandardInput = true;
                procInfo.RedirectStandardOutput = true;
                procInfo.RedirectStandardError = true;
                Log("Running: " + Program.runShellArguments.CMDLine + " " + Program.runShellArguments.CMDLineArguments, true);
                if (bVerboseLog)
                {
                    Console.WriteLine("Running: " + Program.runShellArguments.CMDLine + " " + Program.runShellArguments.CMDLineArguments);
                }
                proc = Process.Start(procInfo);

                System.IO.StreamWriter inputWriter = proc.StandardInput;
                System.IO.StreamReader outputReader = proc.StandardOutput;
                System.IO.StreamReader errorReader = proc.StandardError;

                if (Program.runShellArguments.WaitForThread)
                {
                    while (!proc.HasExited)
                    //Thread.Sleep(100);
                    {
                        Thread.Sleep(100);
                        //Application.DoEvents();
                    }
                    //return proc.ExitCode;
                    Program.runShellArguments.ReturnCode = proc.ExitCode;

                    Program.output = outputReader.ReadToEnd();// +"\r\n" + errorReader.ReadToEnd();
                    //if (bVerboseLog) Console.WriteLine(output);
                    if (Program.output.StartsWith("The following command was not found"))
                    {
                        Log("No wireless interface detected", true);
                        if (bVerboseLog) Console.WriteLine("No wireless interface detected");
                        Program.output = string.Empty;
                    }
                    if (Program.output.StartsWith("The Wireless AutoConfig Service (wlansvc) is not running"))
                    {
                        Log("The Wireless AutoConfig Service (wlansvc) is not running", true);
                        if (bVerboseLog) Console.WriteLine("The Wireless AutoConfig Service (wlansvc) is not running");
                        Program.output = string.Empty;
                    }
                    //Program.output = outputReader;
                }
                else
                    Program.runShellArguments.ReturnCode = -1;
            }
            catch (System.Exception e)
            {
                Console.WriteLine("RunShellAsWorkerThread error: " + e.ToString(), true);
                Program.runShellArguments.ReturnCode = -1;
            }
        }

        /// <summary>
        /// Strictly when copying files from the network .. run only if on corporate  network
        /// </summary>
        /// <param name="netfile"></param>
        /// <param name="localfile"></param>
        /// <returns>
        /// 0 - the copy is successful
        /// 1 - the src and destination are identical (no need to copy)
        /// 2 - the src file is not found
        /// 3 - default error - failed to copy
        /// 4 - not on corp network... only run if on corp network
        /// 5 - no access to the destination location (or destination file cannot be overwritten)
        /// </returns>
        public static FileUpdateResult CopyFromNetIfLocalIsDifferent(string sourceFile, string destinationFile)
        {
            FileUpdateResult functionReturnValue = FileUpdateResult.identical_files; 
            try
            {
                if (sourceFile.StartsWith("\\\\"))
                {
                    currentNetworkContext = GetCurrentNetworkContext();
                    if (currentNetworkContext == NetworkContext.insideIntel)
                    {
                        Log("Checking: " + sourceFile + " vs. " + destinationFile, true);
                        if (System.IO.File.Exists(sourceFile))
                        {
                            if (System.IO.File.Exists(destinationFile))
                            {
                                FileInfo localFile;
                                FileInfo remoteFile;

                                localFile = new FileInfo(destinationFile);
                                remoteFile = new FileInfo(sourceFile);
                                if (!FilesAreEqual(localFile, remoteFile))
                                {
                                    Log("Updating " + destinationFile, true);
                                    if (CopyFileSafe(sourceFile, destinationFile))
                                        functionReturnValue = FileUpdateResult.successful;
                                    else
                                        functionReturnValue = FileUpdateResult.failed_to_copy;
                                }
                            }
                            else
                            {
                                Log("Copy config file to local path: " + destinationFile, true);
                                if (CopyFileSafe(sourceFile, destinationFile))
                                    functionReturnValue = FileUpdateResult.successful;
                                else
                                    functionReturnValue = FileUpdateResult.failed_to_copy;
                            }
                        }
                        else
                        {
                            Log("Cannot find " + sourceFile + " therefore cannot check if the local file can be updated", true);
                            functionReturnValue = FileUpdateResult.src_not_found;
                        }
                    }
                    else
                    {
                        functionReturnValue = FileUpdateResult.not_on_corp_network;
                    }
                }
            }
            catch (Exception ex)
            {
                Log("!Failed to copy from net: " + ex.Message, true);
                functionReturnValue = FileUpdateResult.failed_to_copy;
            }
            return functionReturnValue;
        }


        const int BYTES_TO_READ = sizeof(Int64);
        public static bool FilesAreEqual(FileInfo first, FileInfo second)
        {
            if (first.Length != second.Length)
                return false;

            int iterations = (int)Math.Ceiling((double)first.Length / BYTES_TO_READ);

            using (FileStream fs1 = first.OpenRead())
            using (FileStream fs2 = second.OpenRead())
            {
                byte[] one = new byte[BYTES_TO_READ];
                byte[] two = new byte[BYTES_TO_READ];

                for (int i = 0; i < iterations; i++)
                {
                    fs1.Read(one, 0, BYTES_TO_READ);
                    fs2.Read(two, 0, BYTES_TO_READ);

                    if (BitConverter.ToInt64(one, 0) != BitConverter.ToInt64(two, 0))
                        return false;
                }
            }

            return true;
        }

        public static bool CopyFileSafe(string sourceFile, string destinationFile)
        {
            bool functionReturnValue = false;
            try
            {
                if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(destinationFile)))
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destinationFile));
                if (System.IO.File.Exists(sourceFile))
                {
                    if (System.IO.File.Exists(destinationFile))
                    {
                        //File.SetAttributes(destinationFile, File.GetAttributes(destinationFile) & ~(FileAttributes.ReadOnly));
                        File.SetAttributes(destinationFile, File.GetAttributes(destinationFile) & ~(FileAttributes.ReadOnly) & ~(FileAttributes.Hidden));
                        File.Delete(destinationFile);
                    }
                    System.IO.File.Copy(sourceFile, destinationFile, true);
                    functionReturnValue = true;
                }
                else
                {
                    Console.WriteLine("CopyFileSafe: Cannot find source file(" + sourceFile + ") - check arguments");
                    functionReturnValue = false;
                }
            }
            catch (Exception ex)
            {
                if (bVerboseLog) Console.WriteLine("CopyFileSafe error: " + ex.Message);
                functionReturnValue = false;
            }
            return functionReturnValue;
        }

        #region network

        public static NetworkContext GetCurrentNetworkContext() // ArrayList EnumerateDomains()
        {
            NetworkContext functionReturnValue = NetworkContext.outsideIntel;
            if (!TryPingTheNetwork())
            {
                Debug.Print("Current network context: " + functionReturnValue.ToString());
                functionReturnValue = NetworkContext.outsideIntel;
            }
            else
                functionReturnValue = NetworkContext.insideIntel;


            //if (currentNetworkContext == NetworkContext.insideIntel)
            //{

            //    try
            //    {
            //        Domain currentDomain = Domain.GetCurrentDomain();
            //        DirectoryContext context = new DirectoryContext(DirectoryContextType.Domain, currentDomain.Name); // "amr.corp.intel.com");
            //        foreach (DomainController dc in DomainController.FindAll(context))
            //        {
            //            if (bVerboseLog) Debug.Print("Domain controller: " + dc.Name);
            //            Ping pingDC = new Ping();
            //            PingReply reply = pingDC.Send(dc.Name);
            //            if (reply.Status == IPStatus.Success)
            //            {
            //                functionReturnValue = NetworkContext.insideIntel;
            //                Debug.Print("Ping successful to " + dc.Name);
            //                return functionReturnValue;
            //            }
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        Debug.Print("Unsuccessful attempt to detect corp network - must be outside, or running with an account that is not associated with Active Directory - no reporting will be possible " + ex.Message);
            //        functionReturnValue = NetworkContext.outsideIntel;
            //    }
            //}
            //Debug.Print("Current network context: " + functionReturnValue.ToString());
            return functionReturnValue;
        }

        public static bool TryPingTheNetwork()
        {
            bool functionReturnValue = false;
            currentNetworkContext = NetworkContext.outsideIntel;
            Ping pingSender = new Ping();
            PingOptions options = new PingOptions();
            // Use the default Ttl value which is 128,
            // but change the fragmentation behavior.
            options.DontFragment = true;
            // Create a buffer of 32 bytes of data to be transmitted.
            string data = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
            byte[] buffer = Encoding.ASCII.GetBytes(data);
            int timeout = 1200;
            try
            {
                PingReply reply = pingSender.Send("amr.corp.intel.com", timeout, buffer, options);
                if (reply.Status == IPStatus.Success)
                {
                    currentNetworkContext = NetworkContext.insideIntel;
                    functionReturnValue = true;
                }
                else
                    functionReturnValue = false;
            }
            catch (Exception ex)
            {
                Debug.Print (ex.Message + " Outside corp network?");
                currentNetworkContext = NetworkContext.outsideIntel;
                functionReturnValue = false;
            }
            return functionReturnValue;
        }

        #endregion

        /// <summary>
        /// Information to be written to the log file. 
        /// The parameter writeTimeStampHeader indicates whether the current time of the log should also be written to the log
        /// </summary>
        /// <param name="info"></param>
        /// <param name="writeTimeStampHeader"></param>
        public static void Log(string info, bool writeTimeStampHeader)
        {
            //writeTimeStampHeader == true ? logFile.LogFull(info) : logFile.LogInfo(info);
            if (writeTimeStampHeader)
                Program.logFile.LogFull(info);
            else
                Program.logFile.LogInfo(info);
        }

        /// <summary>
        /// Information to be written to the log file. 
        /// No time stamp will precede the information in the log
        /// </summary>
        /// <param name="info"></param>
        public static void Log(string info)
        {
            Program.logFile.LogInfo(info);
        }



#region WLANInfo - check if wireless can be updated

        public static WLANClientUpdateStatus CheckIfWirelessCanBeUpdated()
        {
            WLANClientUpdateStatus functionReturnValue = WLANClientUpdateStatus.UNKNOWN;

            try
            {
                string ConfigXMLFile = Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\WLANInfoConfig.xml";
                //CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\WLAN\WLANInfoConfig\WLANInfoConfig.xml", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\WLANInfoConfig.xml");

                Log("Current wireless adapter: " + wirelessAdapter, true);

                bool bContinueRunning = true;
                if (wirelessAdapter.ToLower().Equals("unknown") || wirelessAdapter.Equals(string.Empty))
                {
                    Log("Cannot determine the wireles adapter", true);
                    functionReturnValue = WLANClientUpdateStatus.NOT_APPLICABLE;
                    bContinueRunning = false;
                }

                if (bContinueRunning)
                {
                    if (ReadXMLConfig(ConfigXMLFile))
                        Log("Config XML file read OK", true);
                    else
                    {
                        Log("Cannot read config XML file", true);
                        functionReturnValue = WLANClientUpdateStatus.CANNOT_READ_XML_CONFIG;
                        bContinueRunning = false;
                    }
                }

                if (bContinueRunning)
                {
                    /// get the current app version if the informaiton how to get it is provided... 
                    /// 
                    if (wlaninfo.AdapterCheckItem.Type != ChkItem.itemType.unknown)
                    {
                        functionReturnValue = GetCurrentWLANInfo(wlaninfo);
                    }
                    else
                    {
                        if (wlaninfo.AdapterName.Equals(string.Empty))
                            wlaninfo.AdapterName = wirelessAdapter;
                    }
                }
            }
            catch (Exception ex)
            {
                Log("CheckIfWirelessCanBeUpdated: " + ex.Message, true);
                functionReturnValue = WLANClientUpdateStatus.UNKNOWN;
            }
            return functionReturnValue;
        }

        public static WLANClientUpdateStatus GetCurrentWLANInfo(WirelessLANInfo infoSource)
        {
            WLANClientUpdateStatus functionreturnValue = WLANClientUpdateStatus.UNKNOWN;
            wlaninfo.AppVersionFound = GetValueForXMLProperty(wlaninfo);
            if (WirelessIsCurrent(wlaninfo))
            {
                wlaninfo.AdapterCheckItem.Status = ChkItem.itemStatus.CURRENT;
                functionreturnValue = WLANClientUpdateStatus.CURRENT;
            }
            else
            {
                wlaninfo.AdapterCheckItem.Status = ChkItem.itemStatus.NOT_CURRENT;
                functionreturnValue = WLANClientUpdateStatus.UPDATE_AVAILABLE;
            }
            return functionreturnValue;
        }

        public static bool WirelessIsCurrent(WirelessLANInfo infoSource)
        {
            bool functionReturnValue = false;
            try
            {
                /// if left is ON and right is OFF, count he characters from left
                /// if left if OFF and right is ON, count the characters from right
                /// if both left and right are ON, count the characters between left and right .. starting at left ending at right... 
                string stringToCompare = string.Empty;
                if ((infoSource.AdapterCheckItem.Left > 0) & (infoSource.AdapterCheckItem.Right > 0))
                {
                    // noth are non zero... read between left and right
                    stringToCompare = infoSource.AppVersionFound.Substring(infoSource.AdapterCheckItem.Left, infoSource.AdapterCheckItem.Right);
                }
                else
                {
                    if (infoSource.AdapterCheckItem.Left > 0)
                    {
                        stringToCompare = infoSource.AppVersionFound.Substring(0, infoSource.AdapterCheckItem.Left);
                    }
                    else if (infoSource.AdapterCheckItem.Right > 0)
                    {
                        // left is zero, read the last Right characters
                        stringToCompare = infoSource.AppVersionFound.Substring(infoSource.AppVersionFound.Length - infoSource.AdapterCheckItem.Right);
                    }
                    else
                    {
                        // both are zero.. read the entire string
                        stringToCompare = infoSource.AppVersionFound;
                    }
                }
                functionReturnValue = CompareFoundAndCurrent(stringToCompare, infoSource.AppVersionCurrentRelease, infoSource.AdapterCheckItem.ComparisonOperator);
            }
            catch (Exception ex)
            {
                Log("WirelessCurrent error: " + ex.Message, true);
            }

            return functionReturnValue;
        }

        public static bool CompareFoundAndCurrent(string found, string expected, ChkItem.comparisonOperatorType op)
        {
            bool functionReturnValue = false;
            try
            {
                Version verFound = new Version();
                Version verExpected = new Version();
                bool bBothVersionsOK = true;
                if (found.Contains(".") & !found.StartsWith("."))
                    verFound = new Version(found);
                else
                    bBothVersionsOK = false;
                if (expected.Contains(".") & !expected.StartsWith("."))
                    verExpected = new Version(expected);
                else
                    bBothVersionsOK = false;

                switch (op)
                {
                    case ChkItem.comparisonOperatorType.equal:
                        if (bBothVersionsOK)
                        {
                            if (verFound.CompareTo(verExpected) == 0)
                                functionReturnValue = true;
                        }
                        else
                        {
                            if (found.ToLower().Equals(expected.ToLower()))
                                functionReturnValue = true;
                        }
                        break;
                    case ChkItem.comparisonOperatorType.not_equal:
                        if (!found.ToLower().Equals(expected.ToLower()))
                            functionReturnValue = true;
                        break;
                    case ChkItem.comparisonOperatorType.greater_then:
                        if (bBothVersionsOK)
                        {
                            if (verFound.CompareTo(verExpected) > 0)
                                functionReturnValue = true;
                        }
                        else
                        {
                            if (found.CompareTo(expected) > 0)
                                functionReturnValue = true;
                        }
                        break;
                    case ChkItem.comparisonOperatorType.greater_then_or_equal:
                        if (bBothVersionsOK)
                        {
                            if (verFound.CompareTo(verExpected) >= 0)
                                functionReturnValue = true;
                        }
                        else
                        {
                            if (found.CompareTo(expected) >= 0)
                                functionReturnValue = true;
                        }
                        break;
                    case ChkItem.comparisonOperatorType.less_then:
                        if (bBothVersionsOK)
                        {
                            if (verFound.CompareTo(verExpected) < 0)
                                functionReturnValue = true;
                        }
                        else
                        {
                            if (found.CompareTo(expected) < 0)
                                functionReturnValue = true;
                        }
                        break;
                    case ChkItem.comparisonOperatorType.less_then_or_equal:
                        if (bBothVersionsOK)
                        {
                            if (verFound.CompareTo(verExpected) <= 0)
                                functionReturnValue = true;
                        }
                        else
                        {
                            if (found.CompareTo(expected) <= 0)
                                functionReturnValue = true;
                        }
                        break;
                    case ChkItem.comparisonOperatorType.oneof:
                        if (expected.Contains(found))
                            functionReturnValue = true;
                        break;
                    case ChkItem.comparisonOperatorType.substring:
                        if (expected.Contains(found))
                            functionReturnValue = true;
                        break;
                    default: // equal
                        if (found.ToLower().Equals(expected.ToLower()))
                            functionReturnValue = true;
                        break;
                }
            }
            catch (Exception ex)
            {
                Log("CompareFoundAndCurrent error: " + ex.Message, true);
            }
            return functionReturnValue;
        }

        public static string GetValueForXMLProperty(WirelessLANInfo infoSource)
        {
            string functionReturnValue = string.Empty;
            try
            {
                switch (infoSource.AdapterCheckItem.Type)
                {
                    case ChkItem.itemType.unknown:
                        break;
                    case ChkItem.itemType.reg_val:
                        functionReturnValue = GetValueForRegValType(infoSource);
                        break;
                    case ChkItem.itemType.reg_key:
                        functionReturnValue = GetValueForRegKeyType(infoSource);
                        break;
                    case ChkItem.itemType.reg_val_subtree:
                        break;
                    case ChkItem.itemType.file_ver:
                        functionReturnValue = GetValueForFileVerType(infoSource);
                        break;
                    case ChkItem.itemType.app_ver:
                        functionReturnValue = GetValueForAppVerType(infoSource);
                        break;
                    case ChkItem.itemType.drv_ver:
                        functionReturnValue = GetValueForDriverVerType(infoSource);
                        break;
                    case ChkItem.itemType.wlandriver_ver:
                        functionReturnValue = GetValueForWLANDriverVerType(infoSource);
                        break;
                    case ChkItem.itemType.wmi:
                        break;
                    case ChkItem.itemType.service:
                        break;
                    case ChkItem.itemType.os_ver:
                        break;
                    case ChkItem.itemType.sp_ver:
                        break;
                    case ChkItem.itemType.ie_ver:
                        break;
                    case ChkItem.itemType.exact_copy:
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Log("GetvalueForXMLProperty error: " + ex.Message, true);
            }
            return functionReturnValue;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="provider"></param>
        /// <param name="parameter"></param>
        /// <param name="property"></param>
        /// <returns></returns>
        public static string GetValueForRegValType(WirelessLANInfo infoSource)
        {
            string functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
            RegistryKey rk = Registry.LocalMachine;

            try
            {
                switch (infoSource.AdapterCheckItem.Provider.ToLower())
                {
                    case "hklm":
                        rk = Registry.LocalMachine;
                        break;
                    case "hkcu":
                        rk = Registry.CurrentUser;
                        break;
                    case "hkcr":
                        rk = Registry.ClassesRoot;
                        break;
                    case "hku":
                        rk = Registry.Users;
                        break;
                    case "hkpd":
                        rk = Registry.PerformanceData;
                        break;
                    case "hkcc":
                        rk = Registry.CurrentConfig;
                        break;
                    default:
                        rk = Registry.LocalMachine;
                        break;
                }
                RegistryKey rkey = rk.OpenSubKey(infoSource.AdapterCheckItem.Parameter, RegistryKeyPermissionCheck.ReadSubTree);
                if (rkey == null)
                    functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                else
                    functionReturnValue = rkey.GetValue(infoSource.AdapterCheckItem.Property).ToString();
            }
            catch (Exception ex)
            {
                functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                Log("GetvalueForRegValType error: " + ex.Message, true);
            }
            return functionReturnValue;
        }

        public static string GetValueForRegKeyType(WirelessLANInfo infoSource)
        {
            string functionReturnValue = string.Empty;
            RegistryKey rk = Registry.LocalMachine;

            try
            {
                switch (infoSource.AdapterCheckItem.Provider.ToLower())
                {
                    case "hklm":
                        rk = Registry.LocalMachine;
                        break;
                    case "hkcu":
                        rk = Registry.CurrentUser;
                        break;
                    case "hkcr":
                        rk = Registry.ClassesRoot;
                        break;
                    case "hku":
                        rk = Registry.Users;
                        break;
                    case "hkpd":
                        rk = Registry.PerformanceData;
                        break;
                    case "hkcc":
                        rk = Registry.CurrentConfig;
                        break;
                    default:
                        rk = Registry.LocalMachine;
                        break;
                }
                RegistryKey rkKey = rk.OpenSubKey(infoSource.AdapterCheckItem.Parameter);
                if (rkKey == null)
                    functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                else
                    functionReturnValue = infoSource.AdapterCheckItem.TextIfTrue;
            }
            catch (Exception ex)
            {
                functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                Log("GetvalueForRegKeyType error: " + ex.Message, true);
            }
            return functionReturnValue;
        }

        public static string GetValueForWLANDriverVerType(WirelessLANInfo infoSource)
        {
            string functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
            
            try
            {
                if (!wirelessDriverVersion.Equals(string.Empty))
                {
                    functionReturnValue = wirelessDriverVersion;
                }
            }
            catch (Exception ex)
            {
                functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                Log(ex.Message);
            }
            return functionReturnValue;
        }

        public static string GetValueForDriverVerType(WirelessLANInfo infoSource)
        {
            string functionReturnValue = string.Empty;

            string appNameToCheck = string.Empty; // app name 
            try
            {
                RegistryKey rk = Registry.LocalMachine;
                switch (infoSource.AdapterCheckItem.Provider.ToLower())
                {
                    case "deviceid":
                        {
                            RunDevConToGetDevices();
                            string deviceIDPath = @"SYSTEM\CurrentControlSet\Enum\" + wlaninfo.DeviceID;
                            RegistryKey rkKeyDevice = rk.OpenSubKey(deviceIDPath);
                            if (rkKeyDevice == null)
                                functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                            else
                            {
                                string regKeyDriverPath = rkKeyDevice.GetValue(infoSource.AdapterCheckItem.Parameter).ToString();
                                RegistryKey rkKeyDriverPath = rk.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Class\" + regKeyDriverPath);
                                if (rkKeyDriverPath == null)
                                    functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                                else
                                {
                                    functionReturnValue = rkKeyDriverPath.GetValue(infoSource.AdapterCheckItem.Property).ToString();
                                }
                            }


                        }
                        break;
                    default:
                        appNameToCheck = infoSource.AdapterCheckItem.TextIfFalse;
                        break;
                }
            }
            catch (Exception ex)
            {
                functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                Log("GetValueForDriverVerType error: " + ex.Message, true); 
            }
            return functionReturnValue;
        }


        public static string GetValueForAppVerType(WirelessLANInfo infoSource)
        {
            string functionReturnValue = string.Empty;

            string appNameToCheck = string.Empty; // app name 
            try
            {
                RegistryKey rk = Registry.LocalMachine;
                switch (infoSource.AdapterCheckItem.Provider.ToLower())
                {
                    case "appnamesubstring":
                        {
                            foreach (string app in AllInstalledApps)
                            {
                                string[] appInfo = app.Split('|');
                                if (bVerboseLog) Log("- Check if \"" + appInfo[0] + "\" ends with \"" + infoSource.AdapterCheckItem.Parameter + "\"");
                                if (infoSource.AdapterCheckItem.Parameter.Equals("PROSet/Wireless Software"))
                                {
                                    if (appInfo[0].ToLower().EndsWith("pro/wireless driver"))
                                    {
                                        if (infoSource.AdapterCheckItem.Provider.Equals(string.Empty))
                                        /// nothing on the Provider field, indicates only see if the app was install
                                        /// it doesn't matter what version... just check if it is installed
                                        {
                                            functionReturnValue = infoSource.AdapterCheckItem.TextIfTrue;
                                        }
                                        else
                                        {

                                            switch (infoSource.AdapterCheckItem.Property.ToLower())
                                            {
                                                case "version":
                                                    {
                                                        if (appInfo[1].Length > 0)
                                                            functionReturnValue = appInfo[1];
                                                        else
                                                            functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                                                    }
                                                    break;
                                                case "guid":
                                                    {
                                                        if (appInfo[1].Length > 0)
                                                            functionReturnValue = appInfo[1];
                                                        else
                                                            functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                                                    }
                                                    break;
                                                default:
                                                    functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                                                    break;
                                            }
                                        }
                                    }
                                }
                                if (appInfo[0].ToLower().EndsWith(infoSource.AdapterCheckItem.Parameter.ToLower()))
                                {
                                    if (infoSource.AdapterCheckItem.Provider.Equals(string.Empty))
                                    /// nothing on the Provider field, indicates only see if the app was install
                                    /// it doesn't matter what version... just check if it is installed
                                    {
                                        functionReturnValue = infoSource.AdapterCheckItem.TextIfTrue;
                                    }
                                    else
                                    {

                                        switch (infoSource.AdapterCheckItem.Property.ToLower())
                                        {
                                            case "version":
                                                {
                                                    if (appInfo[1].Length > 0)
                                                        functionReturnValue = appInfo[1];
                                                    else
                                                        functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                                                }
                                                break;
                                            case "guid":
                                                {
                                                    if (appInfo[1].Length > 0)
                                                        functionReturnValue = appInfo[1];
                                                    else
                                                        functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                                                }
                                                break;
                                            default:
                                                functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    default:
                        appNameToCheck = infoSource.AdapterCheckItem.TextIfFalse;
                        break;
                }
            }
            catch (Exception ex)
            {
                functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                Log(ex.Message);
            }
            return functionReturnValue;
        }

        public static string GetValueForAppVerTypeOld(WirelessLANInfo infoSource)
        {
            string functionReturnValue = string.Empty;

            string appNameToCheck = string.Empty; // app name 
            try
            {
                RegistryKey rk = Registry.LocalMachine;
                switch (infoSource.AdapterCheckItem.Provider.ToLower())
                {
                    case "appnamesubstring":
                        {
                            foreach (string app in AllInstalledApps)
                            {
                                string[] appInfo = app.Split('|');
                                if (bVerboseLog) Log("Check if \"" + appInfo[0] + "\" ends with \"" + infoSource.AdapterCheckItem.Parameter + "\"");
                                if (appInfo[0].ToLower().EndsWith(infoSource.AdapterCheckItem.Parameter.ToLower())) //if (app.ToLower().Contains(infoSource.AdapterCheckItem.Parameter.ToLower()))
                                {
                                    if (infoSource.AdapterCheckItem.Provider.Equals(string.Empty))
                                    /// nothing on the Provider field, indicates only see if the app was install
                                    /// it doesn't matter what version... just check if it is installed
                                    {
                                        functionReturnValue = infoSource.AdapterCheckItem.TextIfTrue;
                                    }
                                    else
                                    {
                                        
                                        switch (infoSource.AdapterCheckItem.Property.ToLower())
                                        {
                                            case "version":
                                                {
                                                    if (appInfo[1].Length > 0)
                                                        functionReturnValue = appInfo[1];
                                                    else
                                                        functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                                                }
                                                break;
                                            case "guid":
                                                {
                                                    if (appInfo[1].Length > 0)
                                                        functionReturnValue = appInfo[1];
                                                    else
                                                        functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                                                }
                                                break;
                                            default:
                                                functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    default:
                        appNameToCheck = infoSource.AdapterCheckItem.TextIfFalse;
                        break;
                }
            }
            catch (Exception ex)
            {
                functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                Log("GetvalueForAppVerType error: " + ex.Message, true);
            }
            return functionReturnValue;
        }

        public static string GetValueForFileVerType(WirelessLANInfo infoSource)
        {
            string functionReturnValue = string.Empty;
            string fileToCheck = string.Empty;
            try
            {
                RegistryKey rk = Registry.LocalMachine;
                switch (infoSource.AdapterCheckItem.Provider.ToLower())
                {
                    case "hklm":
                        {
                            rk = Registry.LocalMachine;
                            RegistryKey rkKey = rk.OpenSubKey(infoSource.AdapterCheckItem.Parameter);
                            if (rkKey == null)
                                fileToCheck = infoSource.AdapterCheckItem.TextIfFalse;
                            else
                                fileToCheck = rkKey.GetValue(infoSource.AdapterCheckItem.Property).ToString();
                        }
                        break;
                    case "hkcu":
                        {
                            rk = Registry.CurrentUser;
                            RegistryKey rkKey = rk.OpenSubKey(infoSource.AdapterCheckItem.Parameter);
                            if (rkKey == null)
                                fileToCheck = infoSource.AdapterCheckItem.TextIfFalse;
                            else
                                fileToCheck = rkKey.GetValue(infoSource.AdapterCheckItem.Property).ToString();
                        }
                        break;
                    case "path":
                        {
                            if (System.IO.File.Exists(infoSource.AdapterCheckItem.Parameter))
                                fileToCheck = infoSource.AdapterCheckItem.Parameter;
                            else
                            {
                                if (System.IO.File.Exists(infoSource.AdapterCheckItem.Parameter + "\\" + infoSource.AdapterCheckItem.Property))
                                    fileToCheck = infoSource.AdapterCheckItem.Parameter + "\\" + infoSource.AdapterCheckItem.Property;
                            }
                        }
                        break;
                    case "expand_path":
                        {
                            fileToCheck = infoSource.AdapterCheckItem.TextIfFalse;
                            string pathToFile = ExpandPath(infoSource.AdapterCheckItem.Parameter);
                            if (System.IO.File.Exists(pathToFile))
                                fileToCheck = pathToFile;
                            else
                            {
                                if (System.IO.File.Exists(pathToFile + "\\" + infoSource.AdapterCheckItem.Property))
                                    fileToCheck = pathToFile + "\\" + infoSource.AdapterCheckItem.Property;
                            }
                        }
                        break;
                    default:
                        fileToCheck = infoSource.AdapterCheckItem.TextIfFalse;
                        break;
                }

                if (fileToCheck.Equals(infoSource.AdapterCheckItem.TextIfFalse))
                    functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                else
                {
                    Log("Checking file version for " + fileToCheck, true);
                    if (System.IO.File.Exists(fileToCheck))
                    {
                        FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(fileToCheck);
                        functionReturnValue = fvi.FileVersion;
                    }
                    else
                        functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                }

            }
            catch (Exception ex)
            {
                functionReturnValue = infoSource.AdapterCheckItem.TextIfFalse;
                Log("GetvalueForFileVerType error: " + ex.Message, true);
            }
            return functionReturnValue;
        }


        public static void RunDevConToGetDevices()
        {
            try
            {

                string devConCMD = CurrentEXEFolder() + "\\devcon.exe";

                if (System.IO.File.Exists(ProgramDataValue() + "\\Intel\\WLAN\\devcon.exe"))
                    devConCMD = ProgramDataValue() + "\\Intel\\WLAN\\devcon.exe";
                else
                {
                    if (platform == RunningPlatform._32)
                        CopyFromNetIfLocalIsDifferent("\\\\amr.corp.intel.com\\iss\\olwnprod\\SHAREDAPPS\\CCOS\\Tools\\devcon32.exe", ProgramDataValue() + "\\Intel\\WLAN\\devcon.exe");
                    else
                        CopyFromNetIfLocalIsDifferent("\\\\amr.corp.intel.com\\iss\\olwnprod\\SHAREDAPPS\\CCOS\\Tools\\devcon64.exe", ProgramDataValue() + "\\Intel\\WLAN\\devcon.exe");
                    devConCMD = ProgramDataValue() + "\\Intel\\WLAN\\devcon.exe";
                }


                if (!System.IO.Directory.Exists("C:\\Temp"))
                    System.IO.Directory.CreateDirectory("C:\\Temp");

                if (System.IO.File.Exists("C:\\temp\\devices.txt"))
                    File.Delete("C:\\temp\\devices.txt");
                if (System.IO.File.Exists("C:\\temp\\currentdevice.txt"))
                    File.Delete("C:\\temp\\currentdevice.txt");

                string devConFileContent = "@echo off\r\n" +
                    devConCMD + " findall * > C:\\temp\\devices.txt";

                if (!wirelessAdapter.Equals(string.Empty))
                {
                    devConFileContent += "\r\nfindstr /i /c:\"" + wirelessAdapter + "\" C:\\temp\\devices.txt > C:\\temp\\currentdevice.txt";
                    //devConFileContent += "\r\nfindstr /i /c:\"" + wirelessAdapter + "\" C:\\temp\\devices.txt > C:\\temp\\currentdevice.txt";
                }
                CreateCMDFile("c:\\temp\\rundevcon.cmd", devConFileContent);

                //string devConArgs = "findall * > C:\\temp\\devices.txt";

                if (System.IO.File.Exists(devConCMD))
                {
                    //string devConArgs = "findall *";
                    //ThreadParameters getalldevices = new ThreadParameters(devConCMD, devConArgs, ProcessWindowStyle.Hidden, true);
                    //long result = RunShellInWorkerThread(getalldevices);
                    ThreadParameters getalldevices = new ThreadParameters("c:\\temp\\rundevcon.cmd", string.Empty, ProcessWindowStyle.Hidden, true);
                    long result = RunShellInWorkerThread(getalldevices);

                    if (System.IO.File.Exists("C:\\temp\\currentdevice.txt"))
                    {
                        StreamReader sr = new System.IO.StreamReader("C:\\temp\\currentdevice.txt");
                        //string content = sr.ReadToEnd();
                        //sr.Close();

                        //if (content.Contains(":"))
                        //{
                        //    string[] pieces = content.Split(':');
                        //    if (pieces[1].Trim().ToLower().Equals(wirelessAdapter.ToLower()))
                        //        wlaninfo.DeviceID = pieces[0];
                        //}
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.Contains(":"))
                            {
                                string[] pieces = line.Split(':');
                                if (pieces[1].Trim().ToLower().Equals(wirelessAdapter.ToLower()))
                                    wlaninfo.DeviceID = pieces[0];
                            }
                        }
                        sr.Close();
                    }

                }
            }
            catch (Exception ex)
            {
                Log("RunDevConToGetDevices error: " + ex.Message, true);
            }
        }

        public static void RunNetSHToGetWLANReport()
        {
            try
            {
                Version wintver = new Version("10.0");
                if (osVer >= wintver)
                {
                    Log("Running on OS version " + osVer.ToString() + " - will get wlan report info");
                    string netshCMD = Environment.GetFolderPath(Environment.SpecialFolder.System) + "\\netsh.exe";
                    //string netshArgs = "WLAN show drivers"; // > \"" + netshOut + "\"";
                    if (System.IO.File.Exists(netshCMD))
                    {
                        //string devConArgs = "findall *";
                        //ThreadParameters getalldevices = new ThreadParameters(devConCMD, devConArgs, ProcessWindowStyle.Hidden, true);
                        //long result = RunShellInWorkerThread(getalldevices);
                        ThreadParameters netShDrivers = new ThreadParameters(netshCMD, "wlan show wlanreport duration=7", ProcessWindowStyle.Hidden, true);
                        long result = RunShellInWorkerThread(netShDrivers);
                    }
                }
            }
            catch(Exception ex)
            {
                Log("Error running the wlanreport: " + ex.Message, true);
            }
        }

        public static void ConfigureCustomEventLogView()
        {
            try
            {
                if (File.Exists(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\Tools\wireless info.xml"))
                {
                    CopyFileSafe(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\Tools\wireless info.xml", @"C:\ProgramData\Microsoft\Event Viewer\Views\wireless info.xml");
                }
            }
            catch (Exception ex)
            {
                Log("Error updating the custom event log view: " + ex.Message, true);
            }
        }

        public static void RunNetSHToGetDriverVersion()
        {
            try
            {

                string netshCMD = Environment.GetFolderPath(Environment.SpecialFolder.System) + "\\netsh.exe";
                //string netshArgs = "WLAN show drivers"; // > \"" + netshOut + "\"";

                if (!System.IO.Directory.Exists("C:\\Temp"))
                    System.IO.Directory.CreateDirectory("C:\\Temp");

                if (System.IO.File.Exists("C:\\temp\\devices.txt"))
                    File.Delete("C:\\temp\\devices.txt");
                if (System.IO.File.Exists("C:\\temp\\currentdevice.txt"))
                    File.Delete("C:\\temp\\currentdevice.txt");

                string netshCMDFileContent = "@echo off\r\n" +
                    netshCMD + " WLAN show drivers > C:\\temp\\netshdrivers.txt";

                CreateCMDFile("c:\\temp\\runnetsh.cmd", netshCMDFileContent);

                //string devConArgs = "findall * > C:\\temp\\devices.txt";

                if (System.IO.File.Exists(netshCMD))
                {
                    //string devConArgs = "findall *";
                    //ThreadParameters getalldevices = new ThreadParameters(devConCMD, devConArgs, ProcessWindowStyle.Hidden, true);
                    //long result = RunShellInWorkerThread(getalldevices);
                    ThreadParameters netShDrivers = new ThreadParameters("c:\\temp\\runnetsh.cmd", string.Empty, ProcessWindowStyle.Hidden, true);
                    long result = RunShellInWorkerThread(netShDrivers);

                }
            }
            catch (Exception ex)
            {
                Log("RunNETSHtogetDriverversion error: " + ex.Message, true);
            }
        }

        public static void PrepareUpdateAfterTask()
        {
            try
            {

                //CopyFromNetIfLocalIsDifferent(@"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\Tools\UPDLocalFile.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\UPDLocalFile.exe");
                if (currentNetworkContext == NetworkContext.insideIntel)
                {
                    Log("Check if the current task needs to be updated...", true);
                    string updateFile = @"\\amr.corp.intel.com\iss\olwnprod\SHAREDAPPS\CCOS\Tools\UPDLocalFile.exe";
                    if (File.Exists(Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\UPDLocalFile.exe"))
                        updateFile = Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\UPDLocalFile.exe";
                    if (System.IO.File.Exists(updateFile))
                    {
                        string srcFile = "\\\\amr.corp.intel.com\\iss\\olwnprod\\SHAREDAPPS\\CCOS\\WLAN\\Tools\\GetWLANDriversInfo.exe";
                        string destFile = Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\GetWLANDriversInfo.exe";
                        bool bUpdateNow = true;
                        if (File.Exists(srcFile))
                        {
                            if (File.Exists(destFile))
                            {
                                FileVersionInfo srcVI = FileVersionInfo.GetVersionInfo(srcFile);
                                FileVersionInfo destVI = FileVersionInfo.GetVersionInfo(destFile);
                                if (srcVI.FileVersion.Equals(destVI.FileVersion))
                                    bUpdateNow = false;
                                else
                                {
                                    Log("Network GetWLANDriversInfo version: " + srcVI.FileVersion, true);
                                    Log("Local GetWLANDriversInfo version: " + destVI.FileVersion, true);
                                    Log("Prepare update of the current task in 10 seconds", true);
                                }
                            }
                        }

                        if (bUpdateNow)
                        {
                            // "\\\\amr.corp.intel.com\\iss\\olwnprod\\SHAREDAPPS\\CCOS\\WLAN\\Tools\\GetWLANDriversInfo.exe", Environment.GetEnvironmentVariable("SystemDrive") + "\\ProgramData\\Intel\\WLAN\\GetWLANDriversInfo.exe");
                            ThreadParameters updateTask = new ThreadParameters(updateFile, "/src:" + srcFile + " /dest:" + destFile + " /t:10", ProcessWindowStyle.Hidden, false);
                            long result = RunShellInWorkerThread(updateTask);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("PrepareUpdateAfterTask error: " + ex.Message, true);
            }
        }


        //public static void SaveNETSHInfoToRegistry(string output)
        //{

        //    using (System.IO.StringReader reader = new StringReader(output))
        //    {
        //        string line;
        //        bool bFilesLine = false;
        //        //Console.WriteLine(output);
        //        while ((line = reader.ReadLine()) != null)
        //        {
        //            line = line.Trim();
        //            if (line.StartsWith("Interface name"))
        //            {
        //                interfaceName = line.Substring(line.IndexOf(":") + 1).Trim();
        //            }
        //            if (line.StartsWith("Driver"))
        //            {
        //                wirelessAdapter = line.Substring(line.IndexOf(":") + 1).Trim();
        //            }
        //            if (line.StartsWith("Version"))
        //            {
        //                wirelessDriverVersion = line.Substring(line.IndexOf(":") + 1).Trim();
        //                wlaninfo.DriverVersionFound = wirelessDriverVersion;
        //            }
        //            if (bFilesLine)
        //            {
        //                wirelessDriverFile = line.Trim();
        //                bFilesLine = false;
        //            }
        //            if (line.StartsWith("Files"))
        //            {
        //                bFilesLine = true;
        //            }
        //            //Console.WriteLine(line); 
        //        }
        //    }

        //    regKeyPackage = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
        //    if (regKeyPackage == null)
        //    {
        //        Registry.LocalMachine.CreateSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
        //        regKeyPackage = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Intel\\WLAN", RegistryKeyPermissionCheck.ReadWriteSubTree);
        //    }
        //    if (!interfaceName.Equals(string.Empty))
        //    {
        //        regKeyPackage.SetValue("InterfaceName", interfaceName, RegistryValueKind.String);
        //    }
        //    if (!wirelessAdapter.Equals(string.Empty))
        //    {
        //        regKeyPackage.SetValue("WirelessAdapter", wirelessAdapter, RegistryValueKind.String);
        //    }
        //    if (!wirelessDriverFile.Equals(string.Empty))
        //    {
        //        regKeyPackage.SetValue("WirelessDriverFile", wirelessDriverFile, RegistryValueKind.String);
        //    }
        //    if (!wirelessDriverVersion.Equals(string.Empty))
        //    {
        //        regKeyPackage.SetValue("WirelessDriverVersion", wirelessDriverVersion, RegistryValueKind.String);
        //    }
        //}

        private static bool ReadXMLConfig(string xmlfile)
        {
            Log("Reading current configuration file: " + xmlfile, true);
            if (!System.IO.File.Exists(xmlfile))
                return false;

            XmlTextReader reader = new XmlTextReader(xmlfile);

            int iItems = 0;
            try
            {
                XmlDocument doc = new XmlDocument();
                XmlNodeList docNode = default(XmlNodeList);
                doc.Load(reader);

                wlaninfo = new WirelessLANInfo();
                if (Program.bVerboseLog) Log("Prepare to read XML config where adapter[name='" + CleanUpAdapterName(wirelessAdapter.ToLower()) + "']/os[@ver='" + strOSVer + "']", true);
                docNode = doc.DocumentElement.SelectNodes("adapter[name='" + CleanUpAdapterName(wirelessAdapter.ToLower()) + "']/os[@ver='" + strOSVer + "']");
                iItems = docNode.Count;

                foreach (XmlNode node in docNode)
                {
                    if (node.HasChildNodes)
                    {
                        wlaninfo.AdapterName = wirelessAdapter;
                        if (node.InnerXml.Contains("<AppVersionCurrentRelease>")) wlaninfo.AppVersionCurrentRelease = node.SelectSingleNode("AppVersionCurrentRelease").InnerText;
                        if (node.InnerXml.Contains("<provider>")) wlaninfo.AdapterCheckItem.Provider = node.SelectSingleNode("provider").InnerText;
                        if (node.InnerXml.Contains("<parameter>")) wlaninfo.AdapterCheckItem.Parameter = node.SelectSingleNode("parameter").InnerText;
                        if (node.InnerXml.Contains("<property>")) wlaninfo.AdapterCheckItem.Property = node.SelectSingleNode("property").InnerText;
                        if (node.InnerXml.Contains("<chktype>")) wlaninfo.AdapterCheckItem.Type = ChkItem.ItemTypeFromString(node.SelectSingleNode("chktype").InnerText);
                        if (node.InnerXml.Contains("<left>"))
                        {
                            string tmp = node.SelectSingleNode("left").InnerText;
                            short result = 0;
                            if (tmp.Length > 0)
                                Int16.TryParse(tmp, out result);
                            wlaninfo.AdapterCheckItem.Left = result;
                        }
                        if (node.InnerXml.Contains("<right>"))
                        {
                            string tmp = node.SelectSingleNode("right").InnerText;
                            short result = 0;
                            if (tmp.Length > 0)
                                Int16.TryParse(tmp, out result);
                            wlaninfo.AdapterCheckItem.Right = result;
                        }
                        if (node.InnerXml.Contains("<operator>")) wlaninfo.AdapterCheckItem.ComparisonOperator = ChkItem.ComparisonOperatorFromString(node.SelectSingleNode("operator").InnerText);
                        if (node.InnerXml.Contains("<CompareTo>")) wlaninfo.AdapterCheckItem.CompareTO = node.SelectSingleNode("CompareTo").InnerText;
                        if (node.InnerXml.Contains("<textIfTrue>")) wlaninfo.AdapterCheckItem.TextIfTrue = node.SelectSingleNode("textIfTrue").InnerText;
                        if (node.InnerXml.Contains("<textIfFalse>")) wlaninfo.AdapterCheckItem.TextIfFalse = node.SelectSingleNode("textIfFalse").InnerText;
                        if (node.InnerXml.Contains("<path>")) wlaninfo.PathToInstaller = ResolveVariableInString(node.SelectSingleNode("path").InnerText);
                        if (node.InnerXml.Contains("<cmd2run>")) wlaninfo.Installer = node.SelectSingleNode("cmd2run").InnerText;
                        if (node.InnerXml.Contains("<args2cmd>")) wlaninfo.CmdArguments = ResolveVariableInString(node.SelectSingleNode("args2cmd").InnerText);
                        if (node.InnerXml.Contains("<passthrupath>")) wlaninfo.PassThruPath = node.SelectSingleNode("passthrupath").InnerText;
                    }
                }
            }
            catch (Exception ex)
            {
                Log("ReadXMLConfig errro: " + ex.Message, true);
                return false;
            }

            finally
            {
                if ((reader != null))
                {
                    reader.Close();
                }
            }

            return true;
        }

        public static string CleanUpAdapterName(string currentName)
        {
            string functionReturnValue = currentName;
            if (currentName.Contains(" #"))
            {
                functionReturnValue = currentName.Substring(0, currentName.IndexOf(" #"));
            }
            return functionReturnValue;
        }

        public static void CreateCMDFile(string fileName, string cmdFileText)
        {
            try
            {
                if (File.Exists(fileName))
                    File.Delete(fileName);
                File.AppendAllText(fileName, cmdFileText);
            }
            catch (Exception ex)
            {
                Log("CreateCMDFile error: " + ex.Message, true);
            }
        }



	#endregion    

        #region verious toolset methods

        [DllImport("kernel32.Dll")]
        public static extern short GetVersionEx(ref OSVERSIONINFO o);

        [StructLayout(LayoutKind.Sequential)]
        public struct OSVERSIONINFO
        {
            public int dwOSVersionInfoSize;
            public int dwMajorVersion;
            public int dwMinorVersion;
            public int dwBuildNumber;
            public int dwPlatformId;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string szCSDVersion;
        }

        public static string GetOSMajorVersion()
        {
            OSVERSIONINFO os = new OSVERSIONINFO();
            os.dwOSVersionInfoSize = Marshal.SizeOf(typeof(OSVERSIONINFO));
            GetVersionEx(ref os);
            return os.dwMajorVersion.ToString();
        }

        public static int GetOSMajorVersionINT()
        {
            OSVERSIONINFO os = new OSVERSIONINFO();
            os.dwOSVersionInfoSize = Marshal.SizeOf(typeof(OSVERSIONINFO));
            GetVersionEx(ref os);
            return os.dwMajorVersion;
        }

        public static string GetOSMinorVersion()
        {
            OSVERSIONINFO os = new OSVERSIONINFO();
            os.dwOSVersionInfoSize = Marshal.SizeOf(typeof(OSVERSIONINFO));
            GetVersionEx(ref os);
            return os.dwMinorVersion.ToString();
        }

        public static int GetOSMinorVersionINT()
        {
            OSVERSIONINFO os = new OSVERSIONINFO();
            os.dwOSVersionInfoSize = Marshal.SizeOf(typeof(OSVERSIONINFO));
            GetVersionEx(ref os);
            return os.dwMinorVersion;
        }

        public static int GetIntOSMajorVersion()
        {
            OSVERSIONINFO os = new OSVERSIONINFO();
            os.dwOSVersionInfoSize = Marshal.SizeOf(typeof(OSVERSIONINFO));
            GetVersionEx(ref os);
            return os.dwMajorVersion;
        }

        public static int GetIntOSMinorVersion()
        {
            OSVERSIONINFO os = new OSVERSIONINFO();
            os.dwOSVersionInfoSize = Marshal.SizeOf(typeof(OSVERSIONINFO));
            GetVersionEx(ref os);
            return os.dwMinorVersion;
        }

        public static string GetOSFullVersionSTR()
        {
            OSVERSIONINFO os = new OSVERSIONINFO();
            os.dwOSVersionInfoSize = Marshal.SizeOf(typeof(OSVERSIONINFO));
            GetVersionEx(ref os);
            string functionReturnValue = os.dwMajorVersion.ToString() + "." + os.dwMinorVersion.ToString() + "." + os.dwBuildNumber.ToString();
            return functionReturnValue;
        }

        public static Version GetOSFullVersion()
        {
            string ver = Environment.OSVersion.Version.ToString();
            Version currentVer = new Version(ver);
            return currentVer;

            //OSVERSIONINFO os = new OSVERSIONINFO();
            //os.dwOSVersionInfoSize = Marshal.SizeOf(typeof(OSVERSIONINFO));
            //GetVersionEx(ref os);
            //Version functionReturnvalue = new Version(os.dwMajorVersion, os.dwMinorVersion, os.dwBuildNumber);
            //return functionReturnvalue; // os.dwMajorVersion.ToString() + "." + os.dwMinorVersion.ToString() + "." + os.dwBuildNumber.ToString();
        }

        public static Version GetOSVersionFromRegistryOld()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            regKey = rk.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
            string strVer = (string)regKey.GetValue("CurrentVersion", GetOSFullVersionSTR());
            string strProductName = (string)regKey.GetValue("ProductName", "foo");
            if (strProductName.ToLower().Contains("windows 10"))
                strVer = "10.0";
            Version functionReturnValue = new Version(strVer);
            return functionReturnValue;
        }

        public static Version GetOSVersionFromRegistry()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            regKey = rk.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
            string strVer = (string)regKey.GetValue("CurrentVersion", GetOSFullVersionSTR());
            string strMajorVer = regKey.GetValue("CurrentMajorVersionNumber", "Not Set").ToString();
            string strMinorVer = regKey.GetValue("CurrentMinorVersionNumber", "Not Set").ToString();
            string strBuildNumber = regKey.GetValue("CurrentBuildNumber", "Not Set").ToString();
            if (!strMajorVer.Equals("Not Set"))
            {
                strVer = strMajorVer + "." + strMinorVer + "." + strBuildNumber;
            }

            Version functionReturnValue = new Version(strVer);
            return functionReturnValue;
        }



        public static string GetDelimitedText(string text, string openDelimiter, string closeDelimiter, int index)
        {
            int i = 0;
            int j = 0;

            // search the opening mark
            i = text.IndexOf(openDelimiter, index, StringComparison.OrdinalIgnoreCase);
            if (i == -1)
            {
                // no open delimiter found... return an empty string
                index = 0;
                return "";
            }
            i = i + openDelimiter.Length;

            j = text.IndexOf(closeDelimiter, i + 1, StringComparison.OrdinalIgnoreCase);
            if (j == -1)
            {
                // no closing delimiter, therefore return empty string.
                index = 0;
                return "";
            }

            // now get the text between the two indexes i and j
            index = j + closeDelimiter.Length;
            return text.Substring(i, j - i);

        }

        public static string PROSetverFromWLANInfo(WirelessLANInfo wlani)
        {
            return wlani.AppVersionCurrentRelease;
        }


        public static string CurrentEXEFolder()
        {
            return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().FullName);
        }

        public static string PayloadRegPath()
        {
            return @"SOFTWARE\Intel\Payloads\" + Program.friendlyProductName;
        }

        public static string ResolveVariableInString(string val)
        {
            string functionReturnValue = val;
            string variableToBeResolved = GetDelimitedText(val, "%", "%", 0);
            if (val.Contains("%"))
            {
                switch (variableToBeResolved.ToLower())
                {
                    //case "netapppath":
                    //    functionReturnValue = val.Replace("%" + variableToBeResolved + "%", CMTool.CMTMainApp.CMTNetAppPath);
                    //    break;
                    case "args":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", Environment.CommandLine);
                        break;
                    case "toolpath":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", CurrentEXEFolder());
                        if (functionReturnValue.EndsWith("\\"))
                            functionReturnValue.TrimEnd('\\');
                        break;
                    case "prosetver":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", PROSetverFromWLANInfo(Program.wlaninfo));
                        break;
                    case "intelsm":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", intelsm_path);
                        break;
                    case "iss":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", ISS_Path);
                        break;
                    case "isstest":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", ISSTest_Path);
                        break;
                    case "isspath":
                        if (System.IO.Directory.Exists(val.Replace("%" + variableToBeResolved + "%", ISSTest_Path)))
                            functionReturnValue = val.Replace("%" + variableToBeResolved + "%", ISSTest_Path);
                        if (System.IO.Directory.Exists(val.Replace("%" + variableToBeResolved + "%", ISS_Path)))
                            functionReturnValue = val.Replace("%" + variableToBeResolved + "%", ISS_Path);
                        break;
                    case "iss_appid":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", ISS_APPID);
                        break;
                    case "programfiles":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", ProgramFilesValue());
                        break;
                    case "programfiles32":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", ProgramFiles86Value());
                        break;
                    case "programfiles86":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", ProgramFiles86Value());
                        break;
                    case "payloadregpath":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", PayloadRegPath());
                        break;
                    case "payloadname":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", Program.friendlyProductName);
                        break;
                    case "pspath":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", PowerShellPath());
                        break;
                    case "commondesktop":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", CommonDesktop());
                        break;
                    case "commonstartmenu":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", CommonStartMenu());
                        break;
                    case "currentexefolder":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", CurrentEXEFolder());
                        break;
                    case "payloadfilepath":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", PayloadFilePath());
                        break;
                    case "windir":
                    case "systemroot":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", SystemRootValue());
                        break;
                    case "profilepath":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", ProfilePath());
                        break;
                    case "systemdrive":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", Environment.GetEnvironmentVariable("SystemDrive"));
                        break;
                    case "defaultprofilename":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", DefaultUserProfileName());
                        break;
                    case "defaultprofilepath":
                        functionReturnValue = val.Replace("%" + variableToBeResolved + "%", DefaultUserProfilePath());
                        break;
                    default:
                        //LogError("ResolveVariableInString", "Cannot resolve variable " + variableToBeResolved);
                        break;
                }
            }
            else
                functionReturnValue = val;

            return functionReturnValue;
        }

        public static string GetSystem32Folder()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.System);
        }

        public static string PayloadFilePath()
        {
            return CurrentEXEFolder();
        }

        public static string ProgramFiles86Value()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            regKey = rk.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion");
            string functionReturnValue = string.Empty;
            if (Program.platform == RunningPlatform._64)
                functionReturnValue = (string)regKey.GetValue("ProgramFilesDir (x86)", "C:\\Program Files (x86)");
            else
                functionReturnValue = (string)regKey.GetValue("ProgramFilesDir", "C:\\Program Files");

            return functionReturnValue;
        }

        public static string ProgramDataValue()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            regKey = rk.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders");
            string functionReturnValue = (string)regKey.GetValue("Common AppData    ", "C:\\ProgramData");
            return functionReturnValue;
        }

        public static string ProgramFilesValue()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            regKey = rk.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion");
            string functionReturnValue = (string)regKey.GetValue("ProgramFilesDir", "C:\\Program Files");
            return functionReturnValue;
        }

        public static string ProfilePath()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            regKey = rk.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList");
            string defaultProfilePath = Environment.GetEnvironmentVariable("SystemDrive") + "\\Documents and Settings";
            if (GetIntOSMajorVersion() >= 6)
                defaultProfilePath = Environment.GetEnvironmentVariable("SystemDrive") + "\\Users";
            string functionReturnValue = (string)regKey.GetValue("ProfilesDirectory", defaultProfilePath);
            return functionReturnValue;
        }

        public static string DefaultUserProfileName()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            regKey = rk.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList");
            string defaultUserProfile = "Default User";
            if (GetIntOSMajorVersion() >= 6)
                defaultUserProfile = "Default";
            string functionReturnValue = (string)regKey.GetValue("DefaultUserProfile", defaultUserProfile);
            return functionReturnValue;
        }

        public static string DefaultUserProfilePath()
        {
            return ProfilePath() + "\\" + DefaultUserProfileName();
        }

        public static string SystemRootValue()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            regKey = rk.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion");
            string functionReturnValue = (string)regKey.GetValue("SystemRoot", "C:\\WINDOWS");
            return functionReturnValue;
        }

        public string GetEnvVariable(string variableToTranslate)
        {
            return Environment.GetEnvironmentVariable(variableToTranslate);
        }

        public static string ExpandPath(string path)
        {
            string functionReturnValue = null;
            if (!path.Contains("%"))
                return path;

            try
            {
                string sEnvValue = null;
                string sEnvVariable = null;
                int i = 0;
                int j = 0;


                sEnvVariable = GetDelimitedText(path, "%", "%", 0);

                sEnvValue = System.Environment.ExpandEnvironmentVariables("%" + sEnvVariable + "%");

                // search the opening mark 
                i = path.IndexOf("%", 0);
                j = path.IndexOf("%", i + 1);
                if (j == 0)
                {
                    Log("ExpandPath: error - Bad formated environment variable to expand - " + path, true);
                    functionReturnValue = path;
                    return functionReturnValue;
                }

                // get the text between the two Delimiters 
                if (path.Length == j)
                {
                    functionReturnValue = path.Substring(i + 1) + sEnvValue; // ToolSet.Left(path, i - 1) + sEnvValue;
                }
                else
                {
                    functionReturnValue = path.Substring(0, i) + sEnvValue + path.Substring(j + 1);
                }
            }


            catch (System.Exception ex)
            {
                Log("ExpandPath error: " + ex.Message, true);
            }
            return functionReturnValue;
        }


        public static string PowerShellPath()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            // HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell
            regKey = rk.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\PowerShell.exe");
            string psPath = regKey.GetValue("Path", @"%SystemRoot%\system32\WindowsPowerShell\v1.0\").ToString();
            return psPath;
        }

        public static string PowerShellEXE()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            // HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell
            regKey = rk.OpenSubKey(@"SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell");
            string psExe = regKey.GetValue("Path", @"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe").ToString();
            return psExe;
        }

        public static PowerShellExecutionPolicy PowerShellPolicy()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            regKey = rk.OpenSubKey(@"SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell");
            string psPol = regKey.GetValue("ExecutionPolicy", string.Empty).ToString();
            PowerShellExecutionPolicy psep = PowerShellExecutionPolicy.Restricted;
            switch (psPol.ToLower())
            {
                case "allsigned":
                    psep = PowerShellExecutionPolicy.AllSigned;
                    break;
                case "unrestricted":
                    psep = PowerShellExecutionPolicy.Unrestricted;
                    break;
                case "remotesigned":
                    psep = PowerShellExecutionPolicy.RemoteSigned;
                    break;
                default:
                    psep = PowerShellExecutionPolicy.Restricted;
                    break;
            }
            return psep;
        }

        public static string CommonStartMenu()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            regKey = rk.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders");
            string commonSM = regKey.GetValue("Common Start Menu", @"C:\ProgramData\Microsoft\Windows\Start Menu").ToString();
            return commonSM;
        }

        public static string CommonDesktop()
        {
            RegistryKey rk;
            RegistryKey regKey;
            rk = Registry.LocalMachine;
            regKey = rk.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders");
            string commonDesktop = regKey.GetValue("Common Desktop", @"C:\Users\Public\Desktop").ToString();
            return commonDesktop;
        }



        #endregion
    }
}
