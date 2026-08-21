# HelloID-Conn-SA-Full-Exchange-On-Premises-Distribution-Group-Manage-Memberships

| :information_source: Information                                                                                                                                                                                                                                                                                                                                                          |
| :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| This repository contains the connector and configuration code only. The implementer is responsible for acquiring the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements. |

## Description

_HelloID-Conn-SA-Full-Exchange-On-Premises-Distribution-Group-Manage-Memberships_ is a template designed for use with HelloID Service Automation (SA) Delegated Forms. It can be imported into HelloID and customized according to your requirements.

By using this delegated form, you can manage distribution group memberships in Exchange On-Premises. The following options are available:

1.  Search and select the distribution group
2.  View current members of the selected distribution group
3.  Add or remove members from the distribution group
4.  The changes are validated and applied
5.  Distribution group membership updates are processed in Exchange On-Premises

## Getting started

### Requirements

- **Exchange On-Premises PowerShell Access**:<br>
  Access to Exchange On-Premises PowerShell is required. The account used must have sufficient permissions to query distribution groups and manage their memberships.
- **Network Connectivity**:<br>
  The HelloID agent or server must be able to connect to the Exchange On-Premises server via the PowerShell connection URI.
- **PowerShell Remoting Enabled**:<br>
  PowerShell remoting must be enabled on the Exchange server to allow remote management sessions.

### Connection settings

The following user-defined variables are used by the connector.

| Setting               | Description                                               | Mandatory |
| --------------------- | --------------------------------------------------------- | --------- |
| ExchangeConnectionUri | The PowerShell connection URI for Exchange On-Premises    | Yes       |
| ExchangeAdminUsername | The username to connect to Exchange (format: domain\user) | Yes       |
| ExchangeAdminPassword | The password to connect to Exchange                       | Yes       |

## Remarks

### Error Handling and Logging

- **Enhanced Error Messages**: The connector includes detailed error handling with specific line numbers and error context to help troubleshoot issues quickly.
- **Audit Logging**: All operations are logged with appropriate audit information, including successful operations and errors, for tracking and compliance purposes.

### Duplicate Member Detection

- **Automatic Duplicate Handling**: When adding members, the connector automatically detects if a user is already a member of the distribution group and skips the addition without generating an error.

### Session Management

- **Automatic Cleanup**: The connector properly manages PowerShell sessions with automatic cleanup in finally blocks to prevent session leaks.
- **Explicit Command Import**: Only required Exchange commands are imported to minimize memory usage and improve performance.

### Security

- **TLS 1.2 Enforcement**: The connector enforces TLS 1.2 for all connections to ensure secure communication with Exchange On-Premises.
- **Credential Handling**: Credentials are securely converted to PSCredential objects and are never logged or exposed in audit messages.

## Development resources

### API endpoints

The following PowerShell cmdlets are used by the connector:

| Cmdlet                         | Description                               |
| ------------------------------ | ----------------------------------------- |
| Get-DistributionGroup          | Retrieve distribution group information   |
| Get-DistributionGroupMember    | Retrieve members of a distribution group  |
| Add-DistributionGroupMember    | Add a member to a distribution group      |
| Remove-DistributionGroupMember | Remove a member from a distribution group |

### API documentation

- [Exchange PowerShell Documentation](https://learn.microsoft.com/en-us/powershell/exchange/)
- [Connect to Exchange Servers using Remote PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-servers-using-remote-powershell)
- [Get-DistributionGroup](https://learn.microsoft.com/en-us/powershell/module/exchange/get-distributiongroup)
- [Add-DistributionGroupMember](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/add-distributiongroupmember)
- [Remove-DistributionGroupMember](https://learn.microsoft.com/en-us/powershell/module/exchange/remove-distributiongroupmember)

## Getting help

> :bulb: **Tip:**  
> _For more information on Delegated Forms, please refer to our [documentation](https://docs.helloid.com/en/service-automation/delegated-forms.html) pages_.

## HelloID docs

The official HelloID documentation can be found at: https://docs.helloid.com/
