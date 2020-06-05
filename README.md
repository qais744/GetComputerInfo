# GetComputerInfo
C# program for getting information about your Windows PC using WMI/netstat/registry, running this program will retireve the following information and write them to a file:

-Windows Version and Release

-Host Name

-MAC Address

-User Accounts Names (using WMI)

-PC Specs (using WMI)

-Resources Usage (using WMI)

-Running Services and Open Ports (uisng netstat) - from this with some modifications (https://gist.github.com/cheynewallace/5971686)

The file that has the results will be created in "bin/Debug/Results/"

NOTE: Tested on Windows 10
