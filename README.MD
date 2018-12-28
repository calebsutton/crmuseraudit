# CRM User Audit Export

A command-line tool to export user audit data from Microsoft Dynamics Online

## Build

Build with Microsoft Visual Studio 2017 or equivalent

## Requirements
* CommandLineParser 2.3.0
* Microsoft.CrmSdk.CoreAssemblies 9.0.2.4
* Microsoft.CrmSdk.Deployment 9.0.2.4
* Microsoft.CrmSdk.Workflow 9.0.2.4
* Microsoft.CrmSdk.XrmTooling.CoreAssembly 9.0.2.7
* Microsoft.IdentityModel.Clients.ActiveDirectory
* Newtonsoft.Json 6.0.8


## Usage
### Arguments
Running `CRMUserAudit.exe` with no arguments will give you a list of all arguments and descriptions.
```
CRMUserAudit 1.0.0.0
Copyright c  2018
ERROR(S):
Required option 'url' is missing.
Required option 'username' is missing.
Required option 'password' is missing.

  --url               Required. URL of Dynamics 365 instance.

  --username          Required. Username with Audit access.

  --password          Required. Password for user with Audit access.

  --path              (Default: .\) Path to export results.

  --filename          (Default: CRMAuditExport.csv) Filename to export results.

  --days              (Default: 30) Number of days to export data for.

  --filteruser        Username to filter.  If not specified, will export all users except SYSTEM

  --excludeobjects    Logical names of objects to exclude from export.  Separated by commas.

  --help              Display this help screen.

  --version           Display version information.

```
### Examples

--url, --username and --password are the minimum required arguments.  Default values will export all user audit data for the last 30 days.
```
.\CRMUserAudit.exe --url https://contoso.crm.dynamics.com --username caleb@contoso.com --password password
```

--path and --filename allow you to specify where the export file is saved.
```
.\CRMUserAudit.exe --url https://contoso.crm.dynamics.com --username caleb@contoso.com --password password --path "C:\Export\" --filename "Results.csv"
```

--days controls how many days of data will be exported, counting backwards from run time.  A value of 7 will export data for the last 7 days.
```
.\CRMUserAudit.exe --url https://contoso.crm.dynamics.com --username caleb@contoso.com --password password --days 7
```

--filteruser filters the results by a single user, default is to export data for all users except SYSTEM.  Operates on the display name of the user.
```
.\CRMUserAudit.exe --url https://contoso.crm.dynamics.com --username caleb@contoso.com --password password --filteruser "Grace Jackson"
```

--excludeobjects is a comma separated list of objects that you want to exclude from the export.  Operates on the logical name of the object.
```
.\CRMUserAudit.exe --url https://contoso.crm.dynamics.com --username caleb@contoso.com --password password --excludeobjects webuserlogin, webuserevent
```

## Contributing
Please open an issue first to discuss what you would like to change.

## License
MIT - See included LICENSE.MD