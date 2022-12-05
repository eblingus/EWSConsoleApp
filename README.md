# EWSConsoleApp
EWS Client to test connection to EWS service in O365

work in progress

created with VSC

C#


1. Install VSC, add C#
2. Install .NET 48  ndp48-devpack-enu.exe  https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48
3. Use nuget to ???
4. Use dotnet from PS to create Console App
dotnet new console -o EWSConsoleApp3 -f  NET48
5. Install ews-managed-api https://github.com/officedev/ews-managed-api 
dotnet add package Microsoft.Exchange.WebServices --version 2.2.0
6. dotnet add package Azure.Identity



3. nuget manual install:
download from https://www.nuget.org/downloads
save to folder (ex: C:\Program Files (x86)\NuGet\ )
Open Environment Variables, select Path, Edit, New, C:\Program Files (x86)\NuGet\



dotnet build 

dotnet run -- "ClientID" "Secrete" "TenantID"

dotnet publish
