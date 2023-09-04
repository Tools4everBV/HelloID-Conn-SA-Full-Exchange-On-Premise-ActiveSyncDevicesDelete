<!-- Description -->
## Description
This HelloID Service Automation Delegated Form provides the functionality to delete registrered active sync devices in Exchange On-premise

 1. Enter a name to lookup the user account for whom you want to manage devices.
 2. The result will show you a list of matching users. You will need select to correct one
 3. From the available sync devices, you can delete a selected device

## Versioning
| Version | Description | Date |
| - | - | - |
| 1.0.0   | Initial release | 2023/08/17  |

<!-- TABLE OF CONTENTS -->
## Table of Contents
* [Description](#description)
* [All-in-one PowerShell setup script](#all-in-one-powershell-setup-script)
  * [Getting started](#getting-started)
* [Post-setup configuration](#post-setup-configuration)
* [Manual resources](#manual-resources)


## All-in-one PowerShell setup script
The PowerShell script "createform.ps1" contains a complete PowerShell script using the HelloID API to create the complete Form including user defined variables, tasks and data sources.

 _Please note that this script asumes none of the required resources do exists within HelloID. The script does not contain versioning or source control_


### Getting started
Please follow the documentation steps on [HelloID Docs](https://docs.helloid.com/hc/en-us/articles/360017556559-Service-automation-GitHub-resources) in order to setup and run the All-in one Powershell Script in your own environment.


## Post-setup configuration
After the all-in-one PowerShell script has run and created all the required resources. The following items need to be configured according to your own environment
 1. Update the following [user defined variables](https://docs.helloid.com/hc/en-us/articles/360014169933-How-to-Create-and-Manage-User-Defined-Variables)
<table>
  <tr><td><strong>Variable name</strong></td><td><strong>Example value</strong></td><td><strong>Description</strong></td></tr>
  <tr><td>ExchangeConnectionUri</td><td>********</td><td>Exchange server URI</td></tr>
  <tr><td>ExchangeAdminUsername</td><td>domain/user</td><td>Exchange server admin account</td></tr>
  <tr><td>ExchangeAdminPassword</td><td>********</td><td>Exchange server admin password</td></tr>
  <tr><td>ExchangeAuthentication</td><td>kerberos</td><td>Exchange server authentication method</td></tr>
  
</table>

## Manual resources
This Delegated Form uses the following resources in order to run

### Powershell data source '[powershell-datasource]_Delete-activesync-devices-exchange-mailuser-generate-table-wildcard'
This Powershell data source runs a query to search for the user.

### Powershell data source '[powershell-datasource]_Delete-activesync-devices-exchange-mailuser-get-activesync-devices'
This Powershell data source runs a query to search for currently registrered ActiveSyncDevices of the selected user.

### Delegated form task '[task]_Exchange on-premise - Delete ActiveSync Devices'
This delegated form task deletes selected devices.

## Getting help
_If you need help, feel free to ask questions on our [TODO-forum](https://forum.helloid.com/forum/helloid-connectors/service-automation/0000-helloid-sa-exchange-on-premises-activesyncdevices-delete)_

## HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/