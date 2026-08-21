# variables configured in form
$group = $form.gridGroups
$usersToRemove = $form.members.rightToLeft
$usersToAdd = $form.members.leftToRight

# Global variables
# Outcommented as these are set from Global Variables
# $ExchangeConnectionUri = ""
# $ExchangeAdminUsername = ""
# $ExchangeAdminPassword = ""

# PowerShell commands to import
$commands = @(
    "Add-DistributionGroupMember"
    , "Remove-DistributionGroupMember"
)

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

try {
    # Create credentials
    $actionMessage = "creating credentials object"
    
    $securePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $securePassword)
    
    Write-Verbose "Created credentials for user [$ExchangeAdminUsername]"

    # Connect to Exchange On-Premises
    # Docs: https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-servers-using-remote-powershell
    $actionMessage = "connecting to Exchange On-Premises"

    $sessionOptionParams = @{
        SkipCACheck         = $false
        SkipCNCheck         = $false
        SkipRevocationCheck = $false
    }

    $sessionOption = New-PSSessionOption @sessionOptionParams

    $sessionParams = @{
        Authentication    = 'Default'
        ConfigurationName = 'Microsoft.Exchange'
        Credential        = $credential
        ConnectionUri     = $ExchangeConnectionUri
        SessionOption     = $sessionOption
        ErrorAction       = "Stop"
    }

    $exchangeSession = New-PSSession @sessionParams
    $null = Import-PSSession -Session $exchangeSession -DisableNameChecking -AllowClobber -CommandName $commands -ErrorAction Stop

    # Send initial audit log
    $Log = @{
        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premises" # optional (free format text) 
        Message           = "Successfully connected to Exchange using URI [$ExchangeConnectionUri]" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $ExchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$exchangeSession.InstanceId) # optional (free format text) 
    }
    Write-Information -Tags "Audit" -MessageData $log

    if ($usersToAdd.count -gt 0) {
        # Add members
        $actionMessage = "adding users as member to group with displayName [$($group.DisplayName)] and id [$($group.Guid)]"
        foreach ($userToAdd in $usersToAdd) {
            try {
                # Add member to group
                # docs: https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/add-distributiongroupmember?view=exchange-ps
                $actionMessage = "adding user with displayName [$($userToAdd.displayName)] and id [$($userToAdd.Guid)] as member to group with displayName [$($group.displayName)] and id [$($group.Guid)]"
           
                $addGroupMemberSplatParams = @{
                    Identity    = $group.Guid
                    Member      = $userToAdd.Guid                    
                    Confirm     = $false
                    ErrorAction = "Stop"
                    Verbose     = $false
                }
                    
                $addGroupMemberResponse = Add-DistributionGroupMember @addGroupMemberSplatParams

                $Log = @{
                    Action            = "GrantMembership" # optional. ENUM (undefined = default) 
                    System            = "Exchange On-Premises" # optional (free format text) 
                    Message           = "Added user with displayName [$($userToAdd.displayName)] and id [$($userToAdd.Guid)] as member to group with displayName [$($group.displayName)] and id [$($group.Guid)]." # required (free format text) 
                    IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                    TargetDisplayName = $group.DisplayName # optional (free format text) 
                    TargetIdentifier  = $group.Guid # optional (free format text) 
                }
                #send result back  
                Write-Information -Tags "Audit" -MessageData $log
            }
            catch {
                $ex = $PSItem
                if (-not [string]::IsNullOrEmpty($ex.Exception.Data.RemoteException.Message)) {
                    $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Data.RemoteException.Message)"
                    $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Data.RemoteException.Message)"
                }
                else {
                    $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
                    $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
                }

                if ($ex.Exception.Message -like "*already present in the collection*") {
                    # Send auditlog to HelloID    
                    $Log = @{
                        Action            = "GrantMembership" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premises" # optional (free format text) 
                        Message           = "Skipped $($actionMessage). Reason: User is already a member." # required (free format text) 
                        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $group.displayName # optional (free format text)
                        TargetIdentifier  = $group.Guid # optional (free format text)
                    }
                    Write-Information -Tags "Audit" -MessageData $log
                }
                else {
                    # Send auditlog to HelloID   
                    $Log = @{
                        Action            = "GrantMembership" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premises" # optional (free format text) 
                        Message           = $auditMessage # required (free format text) 
                        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $group.displayName # optional (free format text)
                        TargetIdentifier  = $group.Guid # optional (free format text)
                    }
                    Write-Information -Tags "Audit" -MessageData $log
                    Write-Warning $warningMessage
                    Write-Error $auditMessage
                }
            }
        }
    }


    if ($usersToRemove.count -gt 0) {
        # Remove members
        $actionMessage = "removing users as member from group with displayName [$($group.displayName)] and id [$($group.Guid)]"
        foreach ($userToRemove in $usersToRemove) {
            try {
                # Remove member from group
                # docs: https://learn.microsoft.com/en-us/powershell/module/exchange/remove-distributiongroupmember?view=exchange-ps
                $actionMessage = "removing user with displayName [$($userToRemove.displayName)] and id [$($userToRemove.Guid)] as member from group with displayName [$($group.displayName)] and id [$($group.Guid)]"
            
                    
                $removeGroupMemberSplatParams = @{
                    Identity                        = $group.Guid
                    Member                          = $userToRemove.Guid
                    BypassSecurityGroupManagerCheck = $true
                    Confirm                         = $false
                    ErrorAction                     = "Stop"
                    Verbose                         = $false
                }

                $removeGroupMemberResponse = Remove-DistributionGroupMember @removeGroupMemberSplatParams
                        
                $Log = @{
                    Action            = "RevokeMembership" # optional. ENUM (undefined = default) 
                    System            = "Exchange On-Premises" # optional (free format text) 
                    Message           = "Removed user with displayName [$($userToRemove.displayName)] and id [$($userToRemove.Guid)] as member from group with displayName [$($group.displayName)] and id [$($group.Guid)]." # required (free format text) 
                    IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                    TargetDisplayName = $group.DisplayName # optional (free format text) 
                    TargetIdentifier  = $group.Guid # optional (free format text) 
                }
                #send result back  
                Write-Information -Tags "Audit" -MessageData $log                    
            }                           
            catch {
                $ex = $PSItem
                if (-not [string]::IsNullOrEmpty($ex.Exception.Data.RemoteException.Message)) {
                    $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Data.RemoteException.Message)"
                    $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Data.RemoteException.Message)"
                }
                else {
                    $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
                    $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
                }
                if ($ex.Exception.Message -like "*isn't a member of the group*") {
                    # Send auditlog to HelloID   
                    $Log = @{
                        Action            = "RevokeMembership" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premises" # optional (free format text) 
                        Message           = "Skipped $($actionMessage). Reason: User is already no longer a member." # required (free format text) 
                        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $group.displayName # optional (free format text)
                        TargetIdentifier  = $group.Guid # optional (free format text)
                    }
                    Write-Information -Tags "Audit" -MessageData $log
                }                
                else {
                    # Send auditlog to HelloID   
                    $Log = @{
                        Action            = "RevokeMembership" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premises" # optional (free format text) 
                        Message           = $auditMessage # required (free format text) 
                        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $group.displayName # optional (free format text)
                        TargetIdentifier  = $group.Guid # optional (free format text)
                    }
                    Write-Information -Tags "Audit" -MessageData $log
                    Write-Warning $warningMessage
                    Write-Error $auditMessage
                }
            }
        }
    }
}
catch {
    $ex = $PSItem
    if (-not [string]::IsNullOrEmpty($ex.Exception.Message)) {
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
    }
    else {
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception)"
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception)"
    }

    # Send error audit log to HelloID
    $Log = @{
        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premises" # optional (free format text) 
        Message           = $auditMessage # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $mailbox.DisplayName # optional (free format text) 
        TargetIdentifier  = $mailbox.PrimarySmtpAddress # optional (free format text) 
    }
    
    Write-Information -Tags "Audit" -MessageData $log
    Write-Warning $warningMessage
    Write-Error $auditMessage
}
finally {
    # Disconnect from Exchange
    # Docs: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/remove-pssession
    if ($null -ne $exchangeSession) {
        try {
            $deleteExchangeSessionSplatParams = @{
                Session     = $exchangeSession
                Confirm     = $false
                ErrorAction = "Stop"
            }
            $null = Remove-PSSession @deleteExchangeSessionSplatParams

            # Send disconnect audit log
            $Log = @{
                Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premises" # optional (free format text) 
                Message           = "Successfully disconnected from Exchange using URI [$ExchangeConnectionUri]" # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $ExchangeConnectionUri # optional (free format text) 
                TargetIdentifier  = $([string]$exchangeSession.InstanceId) # optional (free format text) 
            }
            Write-Information -Tags "Audit" -MessageData $log
        }
        catch {
            Write-Warning "Failed to disconnect from Exchange using URI [$ExchangeConnectionUri]. Error: $($_.Exception.Message)"
        }
    }
}


