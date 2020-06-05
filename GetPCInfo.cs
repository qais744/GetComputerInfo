using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Diagnostics;
using Microsoft.Win32;
using System.IO;
using System.Net.NetworkInformation;
using System.Management;
using System.Text.RegularExpressions;

namespace ComputerInfo
{
    class Program
    {

        //Getting WMI stuff ready
        static ManagementScope ms = new ManagementScope("\\root\\cimv2");
        static SelectQuery query = new SelectQuery();
        static ManagementObjectSearcher searcher = new ManagementObjectSearcher();
        static ManagementObjectCollection queryCollection = null;
        static ManagementObject wmiObject = new ManagementObject();
        static void Main(string[] args)
        {

            generateComputerInfo();

        }

        public static void generateComputerInfo()
        {


            //Get OS version and Release ID
            var reg = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
            string osVersion = reg.GetValue("ProductName").ToString().Replace(" ", string.Empty).Substring(0, 9);
            string releaseId = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ReleaseId", "").ToString();

            //Get client's MAC address
            string macAddress = NetworkInterface
            .GetAllNetworkInterfaces()
            .Where(nic => nic.OperationalStatus == OperationalStatus.Up && nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
            .Select(nic => nic.GetPhysicalAddress().ToString()).FirstOrDefault();

            //Get host name
            string hostName = Environment.GetEnvironmentVariable("COMPUTERNAME");

            //Get CPU architecture and type
            string cpuArch = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE");

            string cpuType;
            if (Environment.Is64BitOperatingSystem == true)
            {
                cpuType = "_64";
            }
            else
            {
                cpuType = "_32";
            }

            //Get running services and open ports
            List<Port> portsList = new List<Port> { };
            portsList = GetNetStatPorts();

            //Get user accounts information
            List<Users> userAccounts = new List<Users> { };
            userAccounts = getUserAccountsInfo();

            //Get computer specs
            List<Specs> computerSpecs = new List<Specs>();
            computerSpecs = getSpecsAndUsageDetails(1);
            string specs = string.Format("CPU-{0}#RAM-{1}#Storage-{2}", computerSpecs[0].cpuDetails, computerSpecs[0].memoryDetails, computerSpecs[0].diskDetails);

            //Get computer resources usage
            List<Specs> computerUsage = new List<Specs>();
            computerUsage = getSpecsAndUsageDetails(2);
            string usage = string.Format("CPU-{0}#RAM-{1}#Storage-{2}", computerUsage[0].cpuDetails, computerUsage[0].memoryDetails, computerUsage[0].diskDetails);


            //Write the previous stuff to the file
            writeToFile(
                "OS = " + osVersion +
                Environment.NewLine + "Release =" + releaseId +
                Environment.NewLine + "MAC Address = " + macAddress +
                Environment.NewLine + "Hostname = " + hostName +
                Environment.NewLine + "CPU Architecture = " + cpuArch + cpuType);

            //After that, loop through lists and write them to file write ports list with process name
            for (int index = 0; index < portsList.Count; index++)
            {
                if (index == 0)
                {
                    writeToFile(string.Format("Running Services and Open Ports = " + Environment.NewLine + "{0}#{1}#{2}", portsList[index].processName, portsList[index].protocol, portsList[index].portNumber));
                }
                else
                {
                    writeToFile(string.Format("{0}#{1}#{2}", portsList[index].processName, portsList[index].protocol, portsList[index].portNumber));
                }

            }

            //Write user accounts
            for (int index = 0; index < userAccounts.Count; index++)
            {
                if (index == 0)
                {
                    writeToFile(string.Format("User Accounts = " + Environment.NewLine + "{0}", userAccounts[index].userAccount));
                }
                else
                {
                    writeToFile(string.Format("{0}", userAccounts[index].userAccount));
                }
            }

            writeToFile("PC Specs = " + specs);

            writeToFile("Resources Usage = " + usage);

        }

        //Writing to file function
        public static void writeToFile(string Message)
        {
            string path;
            string filePath;

            path = AppDomain.CurrentDomain.BaseDirectory + "\\Results";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            filePath = AppDomain.CurrentDomain.BaseDirectory + "\\Results\\Info_Results.txt";
            if (!File.Exists(filePath))
            {
                using (StreamWriter sw = File.CreateText(filePath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filePath))
                {
                    sw.WriteLine(Message);
                }
            }

        }

        //Function for getting the specs details and usage in a list
        public static List<Specs> getSpecsAndUsageDetails(int info)
        {
            List<Specs> computerSpecsAndUsage = new List<Specs>();
            string cpu = null;
            string memory = null;
            string disk = null;

            //Number 1 for getting the PC's spec ,, number 2 for getting the PC's resources usage
            switch (info)
            {

                case 1:
                    {
                        //Baby, select the required info from WMI classes
                        query = new SelectQuery("SELECT Name FROM Win32_Processor");
                        searcher = new ManagementObjectSearcher(ms, query);
                        queryCollection = searcher.Get();
                        wmiObject = queryCollection.OfType<ManagementObject>().FirstOrDefault();
                        cpu = wmiObject["Name"].ToString().Replace(" ", string.Empty);

                        long memSize = 0;
                        long memCap = 0;
                        query = new SelectQuery("SELECT Capacity FROM Win32_PhysicalMemory");
                        searcher = new ManagementObjectSearcher(ms, query);
                        queryCollection = searcher.Get();
                        wmiObject = queryCollection.OfType<ManagementObject>().FirstOrDefault();
                        memCap = Convert.ToInt64(wmiObject["Capacity"]);
                        memSize += memCap;
                        memSize = (memSize / 1024);
                        memory = memSize.ToString();

                        long diskSize = 0;
                        query = new SelectQuery("SELECT Size FROM Win32_DiskDrive");
                        searcher = new ManagementObjectSearcher(ms, query);
                        queryCollection = searcher.Get();
                        wmiObject = queryCollection.OfType<ManagementObject>().FirstOrDefault();
                        diskSize = Convert.ToInt64(wmiObject["Size"]);
                        disk = (diskSize / 1024).ToString();

                        computerSpecsAndUsage?.Add(new Specs
                        {
                            cpuDetails = cpu,
                            memoryDetails = memory,
                            diskDetails = disk
                        });

                        break;
                    }

                case 2:
                    {

                        //Babe, select the required info from WMI classes
                        query = new SelectQuery("SELECT LoadPercentage FROM Win32_Processor");
                        searcher = new ManagementObjectSearcher(ms, query);
                        queryCollection = searcher.Get();
                        wmiObject = queryCollection.OfType<ManagementObject>().FirstOrDefault();
                        cpu = wmiObject["LoadPercentage"].ToString();

                        query = new SelectQuery("SELECT FreePhysicalMemory FROM Win32_OperatingSystem");
                        searcher = new ManagementObjectSearcher(ms, query);
                        queryCollection = searcher.Get();
                        wmiObject = queryCollection.OfType<ManagementObject>().FirstOrDefault();
                        memory = wmiObject["FreePhysicalMemory"].ToString();

                        long diskFreeSapce;
                        query = new SelectQuery("SELECT FreeSpace FROM Win32_LogicalDisk");
                        searcher = new ManagementObjectSearcher(ms, query);
                        queryCollection = searcher.Get();
                        wmiObject = queryCollection.OfType<ManagementObject>().FirstOrDefault();
                        diskFreeSapce = Convert.ToInt64(wmiObject["FreeSpace"]);
                        disk = (diskFreeSapce / 1024).ToString();

                        computerSpecsAndUsage?.Add(new Specs
                        {
                            cpuDetails = cpu,
                            memoryDetails = memory,
                            diskDetails = disk
                        });

                        break;
                    }

            }

            return computerSpecsAndUsage;
        }

        //Function for getting the user accounts details in a list
        public static List<Users> getUserAccountsInfo()
        {
            var Users = new List<Users>();
            query = new SelectQuery("SELECT Name FROM Win32_UserAccount");
            searcher = new ManagementObjectSearcher(ms, query);

            foreach (ManagementObject envVar in searcher.Get())
            {
                Users?.Add(new Users
                {
                    userAccount = envVar["Name"]?.ToString(),
                });
            }
            return Users;
        }

        //Function for getting the open ports with the running services
        public static List<Port> GetNetStatPorts()
        {
            var Ports = new List<Port>();

            using (Process p = new Process())
            {
                //Powershell command to run netstat
                ProcessStartInfo ps = new ProcessStartInfo();
                ps.Arguments = "-a -n -o";
                ps.FileName = "netstat.exe";
                ps.UseShellExecute = false;
                ps.WindowStyle = ProcessWindowStyle.Hidden;
                ps.RedirectStandardInput = true;
                ps.RedirectStandardOutput = true;
                ps.RedirectStandardError = true;

                p.StartInfo = ps;
                p.Start();

                StreamReader stdOutput = p.StandardOutput;
                StreamReader stdError = p.StandardError;

                string content = stdOutput.ReadToEnd() + stdError.ReadToEnd();
                string exitStatus = p.ExitCode.ToString();

                if (exitStatus != "0")
                {
                    // Command errored
                }

                //Get the Rows
                string[] rows = Regex.Split(content, "\r\n");
                for (int index = 0; index < rows.Length; index++)
                {
                    //Split it babe
                    string[] tokens = Regex.Split(rows[index], "\\s+");
                    if (tokens.Length > 4 && (tokens[1].Equals("UDP") || tokens[1].Equals("TCP")))
                    {
                        //Save the open ports and the services in a list
                        string localAddress = Regex.Replace(tokens[2], @"\[(.*?)\]", "1.1.1.1");
                        Ports.Add(new Port
                        {
                            protocol = localAddress.Contains("1.1.1.1") ? string.Format("{0}v6", tokens[1]) : string.Format("{0}v4", tokens[1]),
                            portNumber = localAddress.Split(':')[1],
                            processName = tokens[1] == "UDP" ? lookupProcess(Convert.ToInt16(tokens[4])) : lookupProcess(Convert.ToInt16(tokens[5]))
                        });
                    }
                }

            }

            return Ports;
        }

        //A function to find the process name via its ID (this function is used to find the service running on a port)
        public static string lookupProcess(int pid)
        {
            string procName;
            try { procName = Process.GetProcessById(pid).ProcessName; }
            catch (Exception) { procName = "-"; }
            return procName;
        }

        //Specs struct that we created a list from (used for specs and usage)
        public struct Specs
        {
            public string cpuDetails;
            public string memoryDetails;
            public string diskDetails;
        }

        //Port struct that we created the a list from (used for getting the open ports and running services)
        public struct Port
        {
            public string portNumber;
            public string processName;
            public string protocol;
        }

        //Users struct that we created the list from (used for getting the users information)
        public struct Users
        {
            public string userAccount;
        }


    }
}



