$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form
$groupName = $form.gridGroups.userPrincipalName
$groupDisplayName = $form.gridGroups.displayName
$usersToRemove = $form.members.rightToLeft
$usersToAdd = $form.members.leftToRight

try {
    <#----- Exchange On-Premises: Start -----#>
    # Connect to Exchange
    try {
        $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
        $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername,$adminSecurePassword)
        $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop 
        #-AllowRedirection
        $session = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber

        Write-Information "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" 
    
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }
    catch {
        Write-Error "Error connecting to Exchange using the URI [$exchangeConnectionUri]. Error: $($_.Exception.Message)"
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Failed to connect to Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }

    if ($usersToAdd.count -gt 0) {
        try {
            Write-Information "Starting to add distribution group [$groupName] to users [$($usersToAdd.sAMAccountName)]"
            #$usersToAddJson = $usersToAdd | ConvertFrom-Json        
            foreach ($user in $usersToAdd) {
                try {
                    Add-DistributionGroupMember -Identity $groupName -Member $user.sAMAccountName -Confirm:$false -ErrorAction Stop
                    Write-Information "Finished adding $($user.sAMAccountName) to distribution group [$groupName]"
                    $Log = @{
                        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Successfully added [$($user.sAMAccountName)] to distribution group [$groupName]" # required (free format text) 
                        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $($user.name) # optional (free format text) 
                        TargetIdentifier  = $groupDisplayName # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log       
                }
                catch {
                    Write-Error "Error adding $($user.name) to distribution group [$groupName]. Error: $($_.Exception.Message)" 
                    $Log = @{
                        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Failed to add [$($user.sAMAccountName)] to distribution group [$groupName]" # required (free format text) 
                        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $($user.name) # optional (free format text) 
                        TargetIdentifier  = $groupDisplayName # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log                    
                }                         
            }                        
        }
        catch {
            Write-Error "Could not add distribution group [$groupName] to users [$($usersToAdd.sAMAccountName)]. Error: $($_.Exception.Message)"
            $Log = @{
                Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premise" # optional (free format text) 
                Message           = "Failed to add distribution group [$groupName] to users [$($usersToAdd.sAMAccountName)]" # required (free format text) 
                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $($usersToAdd.name) # optional (free format text) 
                TargetIdentifier  = $groupDisplayName # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log            
        }
    }


    if ($usersToRemove.count -gt 0) {
        try {
            Write-Information "Starting to remove distribution group [$groupName] from users [$($usersToRemove.sAMAccountName)]"
            #$usersToRemoveJson = $usersToRemove | ConvertFrom-Json            
            foreach ($user in $usersToRemove) {
                try {
                    Remove-DistributionGroupMember -Identity $groupName -Member $user.sAMAccountName -Confirm:$false -ErrorAction Stop
                    Write-Information "Finished removing  [$($user.name)] from distribution group [$groupName]" 
                    $Log = @{
                        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Successfully removed [$($user.sAMAccountName)] from distribution group [$groupName]" # required (free format text) 
                        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $($user.name) # optional (free format text) 
                        TargetIdentifier  = $groupDisplayName # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log                    
                }
                catch {
                    Write-Error "Failed to remove  [$($user.sAMAccountName)] from distribution group [$groupName]. Error: $($_.Exception.Message)" 
                    $Log = @{
                        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Failed to remove  [$($user.sAMAccountName)] from distribution group [$groupName]" # required (free format text) 
                        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $($user.name) # optional (free format text) 
                        TargetIdentifier  = $groupDisplayName # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log 
                }
            }
        }
        catch {
            Write-Error "Could not remove distribution group [$groupName] from users [$($usersToRemove.sAMAccountName)] Error: $($_.Exception.Message)"            
            $Log = @{
                Action            = "UpdateResource" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premise" # optional (free format text) 
                Message           = "Failed to remove distribution group [$groupName] from users [$($usersToRemove.sAMAccountName)]" # required (free format text) 
                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $($usersToRemove.name) # optional (free format text) 
                TargetIdentifier  = $groupDisplayName # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log            
        }    
    }
}
catch {
    Write-Error "Could not set memberships on distribution group [$($groupName)]. Error: $($_.Exception.Message)"    
    $Log = @{
        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Failed setting memberships on distribution group [$groupName]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $groupDisplayName # optional (free format text) 
        TargetIdentifier  = $groupName # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
finally {
    # Disconnect from Exchange
    try {
        Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
        Write-Information "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]"     
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log
    }
    catch {
        Write-Error "Error disconnecting from Exchange.  Error: $($_.Exception.Message)"
        $Log = @{
            Action            = "UpdateResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Failed to disconnect from Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log 
    }
    <#----- Exchange On-Premises: End -----#>
}


