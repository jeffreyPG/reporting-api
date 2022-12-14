sl C:\reports\ReportsAPI\reports

# Restore the nuget references
dotnet restore

# Publish application with all of its dependencies and runtime for IIS to use
dotnet publish --configuration release -o c:\reports\publish --runtime active


# Point IIS wwwroot of the published folder. CodeDeploy uses 32 bit version of PowerShell.
# To make use the IIS PowerShell CmdLets we need call the 64 bit version of PowerShell.
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -Command {Import-Module WebAdministration; Set-ItemProperty 'IIS:\sites\Default Web Site' -Name physicalPath -Value c:\reports\publish}