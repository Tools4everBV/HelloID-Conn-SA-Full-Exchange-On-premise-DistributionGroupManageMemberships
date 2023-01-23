<#----- Exchange On-Premises: Start -----#>
# Connect to Exchange
try{
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername,$adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck #-SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Default -ErrorAction Stop 
    #-AllowRedirection
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber
    Write-Information -Message "Successfully connected to Exchange using the URI [$exchangeConnectionUri]"
} catch {
    Write-Information -Message "Error connecting to Exchange using the URI [$exchangeConnectionUri]"
    Write-Error -Message "Error at line: $($_.InvocationInfo.ScriptLineNumber - 79): $($_.Exception.Message)"
    if($debug -eq $true){
        Write-Error -Message "$($_.Exception)"
    }
    Write-Information -Message "Failed to connect to Exchange using the URI [$exchangeConnectionUri]"
    throw $_
}

try {
        
        $mailUsers = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize:Unlimited | Select-Object DisplayName, SamAccountName
        $mailContacts = Get-MailContact  | Select-Object @{N='DisplayName';E={$_.DisplayName + " (contact)"}}, @{N='SamAccountName';E={$_.Alias}}

        $mailboxes = $mailUsers + $mailContacts

        $mailboxes = $mailboxes | Sort-Object -Property DisplayName
        $resultCount = @($mailboxes).Count
        Write-Information "Result count: $resultCount"
        if($resultCount -gt 0)
        {
            foreach($mailbox in $mailboxes){
                $returnObject = @{
                    name=$mailbox.DisplayName  + " [" + $mailbox.SamAccountName + "]"; 
                    sAMAccountName=$mailbox.SamAccountName
                }
                Write-Output $returnObject
            }
        }

        <#$mailContacts = Get-MailContact  | Select-Object DisplayName, SamAccountName

        $mailContacts = $mailContacts | Sort-Object -Property DisplayName
        $resultCount = @($mailContacts).Count
        Write-Information "Result count: $resultCount"
        if($resultCount -gt 0)
        {
            foreach($mailContact in $mailContacts){
                $returnObject = @{
                    name=$mailContact.DisplayName  + " [" + $mailContact.Alias + " - contact]"; 
                    sAMAccountName=$mailContact.Alias
                }
                Write-Output $returnObject
            }
        }#>
    }
  catch {
    $msg = "Error searching AD user [$searchValue]. Error: $($_.Exception.Message)"
    Write-Error $msg
}

# Disconnect from Exchange
try{
    Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
    Write-Information -Message "Successfully disconnected from Exchange"
} catch {
    Write-Error -Message "Error disconnecting from Exchange"
    Write-Error -Message "Error at line: $($_.InvocationInfo.ScriptLineNumber - 79): $($_.Exception.Message)"
    if($debug -eq $true){
        Write-Error -Message "$($_.Exception)"
    }    
    Write-Error -Message "Failed to disconnect from Exchange"
    throw $_
}
<#----- Exchange On-Premises: End -----#>
