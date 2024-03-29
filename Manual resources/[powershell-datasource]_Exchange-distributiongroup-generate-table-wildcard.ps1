<#----- Exchange On-Premises: Start -----#>
# Connect to Exchange
try{
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername,$adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
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
    $searchValue = $dataSource.searchValue
    $searchQuery = "*$searchValue*"
    
    
    if([String]::IsNullOrEmpty($searchValue) -eq $true){
        Write-Information "Geen Searchvalue"
        return
    }else{
        Write-Information "SearchQuery: $searchQuery"
        
        $distributionGroups = Get-DistributionGroup -filter "{alias -like '$searchQuery' -or name -like '$searchQuery'}" 

        $distributionGroups = $distributionGroups | Sort-Object -Property DisplayName
        $resultCount = @($distributionGroups).Count
        Write-Information "Result count: $resultCount"
        if($resultCount -gt 0){
            foreach($distributionGroup in $distributionGroups){
                $returnObject = @{displayName=$distributionGroup.DisplayName; UserPrincipalName=$distributionGroup.PrimarySMTPAddress}
                Write-Output $returnObject
            }
        }
    }
} catch {
    $msg = "Error searching distribution group [$searchValue]. Error: $($_.Exception.Message)"
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
