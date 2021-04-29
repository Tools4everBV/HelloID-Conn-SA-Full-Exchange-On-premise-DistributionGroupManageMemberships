try {
    <#----- Exchange On-Premises: Start -----#>
    # Connect to Exchange
    try {
        $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
        $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername, $adminSecurePassword
        $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop 
        #-AllowRedirection
        $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber
        HID-Write-Status -Message "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" -Event Success
    }
    catch {
        HID-Write-Status -Message "Error connecting to Exchange using the URI [$exchangeConnectionUri]" -Event Error
        HID-Write-Status -Message "Error at line: $($_.InvocationInfo.ScriptLineNumber - 79): $($_.Exception.Message)" -Event Error
        if ($debug -eq $true) {
            HID-Write-Status -Message "$($_.Exception)" -Event Error
        }
        HID-Write-Summary -Message "Failed to connect to Exchange using the URI [$exchangeConnectionUri]" -Event Failed
        throw $_
    }

    if ($usersToAdd -ne "[]") {
        try {
            HID-Write-Status -Message "Starting to add distribution group [$groupName] to users $usersToAdd" -Event Information
            $usersToAddJson = $usersToAdd | ConvertFrom-Json        
            foreach ($user in $usersToAddJson) {
                Add-DistributionGroupMember -Identity $groupName -Member $user.sAMAccountName -Confirm:$false -ErrorAction Stop
                HID-Write-Status -Message "Finished adding $($user.name) to distribution group [$groupName]" -Event Success
                HID-Write-Summary -Message "Successfully added $($user.name) to distribution group [$groupName]" -Event Success
            }
        }
        catch {
            HID-Write-Status -Message "Could not add distribution group [$groupName] to users $usersToAdd. Error: $($_.Exception.Message)" -Event Error
            HID-Write-Summary -Message "Failed to add distribution group [$groupName] to users $usersToAdd" -Event Failed
        }
    }


    if ($usersToRemove -ne "[]") {
        try {
            HID-Write-Status -Message "Starting to remove distribution group [$groupName] from users $usersToRemove" -Event Information
            $usersToRemoveJson = $usersToRemove | ConvertFrom-Json            
            foreach ($user in $usersToRemoveJson) {
                Remove-DistributionGroupMember -Identity $groupName -Member $user.sAMAccountName -Confirm:$false -ErrorAction Stop
                HID-Write-Status -Message "Finished removing  $($user.name) from distribution group [$groupName]" -Event Success
                HID-Write-Summary -Message "Successfully removed  $($user.name) from distribution group [$groupName]" -Event Success
            }
        }
        catch {
            HID-Write-Status -Message "Could not remove distribution group [$groupName] from users $usersToRemove. Error: $($_.Exception.Message)" -Event Error
            HID-Write-Summary -Message "Failed to remove distribution group [$groupName] from users $usersToRemove" -Event Failed
        }    
    }
}
catch {
    HID-Write-Status -Message "Error removing access rights for distribution group [$($groupName)] to the user [$($user.sAMAccountName)]. Error: $($_.Exception.Message)" -Event Error
    HID-Write-Summary -Message "Error removing access rights for distribution group [$($groupName)] to the user [$($user.sAMAccountName)]" -Event Failed
}
finally {
    # Disconnect from Exchange
    try {
        Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
        HID-Write-Status -Message "Successfully disconnected from Exchange" -Event Success
    }
    catch {
        HID-Write-Status -Message "Error disconnecting from Exchange" -Event Error
        HID-Write-Status -Message "Error at line: $($_.InvocationInfo.ScriptLineNumber - 79): $($_.Exception.Message)" -Event Error
        if ($debug -eq $true) {
            HID-Write-Status -Message "$($_.Exception)" -Event Error
        }
        HID-Write-Summary -Message "Failed to disconnect from Exchange" -Event Failed
        throw $_
    }
    <#----- Exchange On-Premises: End -----#>
}


