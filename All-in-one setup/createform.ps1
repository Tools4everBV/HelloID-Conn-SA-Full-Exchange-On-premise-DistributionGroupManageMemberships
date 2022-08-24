# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

#HelloID variables
#Note: when running this script inside HelloID; portalUrl and API credentials are provided automatically (generate and save API credentials first in your admin panel!)
$portalUrl = "https://CUSTOMER.helloid.com"
$apiKey = "API_KEY"
$apiSecret = "API_SECRET"
$delegatedFormAccessGroupNames = @("Users") #Only unique names are supported. Groups must exist!
$delegatedFormCategories = @("Exchange On-Premise") #Only unique names are supported. Categories will be created if not exists
$script:debugLogging = $false #Default value: $false. If $true, the HelloID resource GUIDs will be shown in the logging
$script:duplicateForm = $false #Default value: $false. If $true, the HelloID resource names will be changed to import a duplicate Form
$script:duplicateFormSuffix = "_tmp" #the suffix will be added to all HelloID resource names to generate a duplicate form with different resource names

#The following HelloID Global variables are used by this form. No existing HelloID global variables will be overriden only new ones are created.
#NOTE: You can also update the HelloID Global variable values afterwards in the HelloID Admin Portal: https://<CUSTOMER>.helloid.com/admin/variablelibrary
$globalHelloIDVariables = [System.Collections.Generic.List[object]]@();

#Global variable #1 >> ExchangeConnectionUri
$tmpName = @'
ExchangeConnectionUri
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #2 >> ExchangeAdminPassword
$tmpName = @'
ExchangeAdminPassword
'@ 
$tmpValue = ""  
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #3 >> ExchangeAdminUsername
$tmpName = @'
ExchangeAdminUsername
'@ 
$tmpValue = ""  
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});


#make sure write-information logging is visual
$InformationPreference = "continue"

# Check for prefilled API Authorization header
if (-not [string]::IsNullOrEmpty($portalApiBasic)) {
    $script:headers = @{"authorization" = $portalApiBasic}
    Write-Information "Using prefilled API credentials"
} else {
    # Create authorization headers with HelloID API key
    $pair = "$apiKey" + ":" + "$apiSecret"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $key = "Basic $base64"
    $script:headers = @{"authorization" = $Key}
    Write-Information "Using manual API credentials"
}

# Check for prefilled PortalBaseURL
if (-not [string]::IsNullOrEmpty($portalBaseUrl)) {
    $script:PortalBaseUrl = $portalBaseUrl
    Write-Information "Using prefilled PortalURL: $script:PortalBaseUrl"
} else {
    $script:PortalBaseUrl = $portalUrl
    Write-Information "Using manual PortalURL: $script:PortalBaseUrl"
}

# Define specific endpoint URI
$script:PortalBaseUrl = $script:PortalBaseUrl.trim("/") + "/"  

# Make sure to reveive an empty array using PowerShell Core
function ConvertFrom-Json-WithEmptyArray([string]$jsonString) {
    # Running in PowerShell Core?
    if($IsCoreCLR -eq $true){
        $r = [Object[]]($jsonString | ConvertFrom-Json -NoEnumerate)
        return ,$r  # Force return value to be an array using a comma
    } else {
        $r = [Object[]]($jsonString | ConvertFrom-Json)
        return ,$r  # Force return value to be an array using a comma
    }
}

function Invoke-HelloIDGlobalVariable {
    param(
        [parameter(Mandatory)][String]$Name,
        [parameter(Mandatory)][String][AllowEmptyString()]$Value,
        [parameter(Mandatory)][String]$Secret
    )

    $Name = $Name + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl + "api/v1/automation/variables/named/$Name")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
    
        if ([string]::IsNullOrEmpty($response.automationVariableGuid)) {
            #Create Variable
            $body = @{
                name     = $Name;
                value    = $Value;
                secret   = $Secret;
                ItemType = 0;
            }    
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl + "api/v1/automation/variable")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $variableGuid = $response.automationVariableGuid

            Write-Information "Variable '$Name' created$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        } else {
            $variableGuid = $response.automationVariableGuid
            Write-Warning "Variable '$Name' already exists$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        }
    } catch {
        Write-Error "Variable '$Name', message: $_"
    }
}

function Invoke-HelloIDAutomationTask {
    param(
        [parameter(Mandatory)][String]$TaskName,
        [parameter(Mandatory)][String]$UseTemplate,
        [parameter(Mandatory)][String]$AutomationContainer,
        [parameter(Mandatory)][String][AllowEmptyString()]$Variables,
        [parameter(Mandatory)][String]$PowershellScript,
        [parameter()][String][AllowEmptyString()]$ObjectGuid,
        [parameter()][String][AllowEmptyString()]$ForceCreateTask,
        [parameter(Mandatory)][Ref]$returnObject
    )
    
    $TaskName = $TaskName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl +"api/v1/automationtasks?search=$TaskName&container=$AutomationContainer")
        $responseRaw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false) 
        $response = $responseRaw | Where-Object -filter {$_.name -eq $TaskName}
    
        if([string]::IsNullOrEmpty($response.automationTaskGuid) -or $ForceCreateTask -eq $true) {
            #Create Task

            $body = @{
                name                = $TaskName;
                useTemplate         = $UseTemplate;
                powerShellScript    = $PowershellScript;
                automationContainer = $AutomationContainer;
                objectGuid          = $ObjectGuid;
                variables           = (ConvertFrom-Json-WithEmptyArray($Variables));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/automationtasks/powershell")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $taskGuid = $response.automationTaskGuid

            Write-Information "Powershell task '$TaskName' created$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        } else {
            #Get TaskGUID
            $taskGuid = $response.automationTaskGuid
            Write-Warning "Powershell task '$TaskName' already exists$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        }
    } catch {
        Write-Error "Powershell task '$TaskName', message: $_"
    }

    $returnObject.Value = $taskGuid
}

function Invoke-HelloIDDatasource {
    param(
        [parameter(Mandatory)][String]$DatasourceName,
        [parameter(Mandatory)][String]$DatasourceType,
        [parameter(Mandatory)][String][AllowEmptyString()]$DatasourceModel,
        [parameter()][String][AllowEmptyString()]$DatasourceStaticValue,
        [parameter()][String][AllowEmptyString()]$DatasourcePsScript,        
        [parameter()][String][AllowEmptyString()]$DatasourceInput,
        [parameter()][String][AllowEmptyString()]$AutomationTaskGuid,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $DatasourceName = $DatasourceName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    $datasourceTypeName = switch($DatasourceType) { 
        "1" { "Native data source"; break} 
        "2" { "Static data source"; break} 
        "3" { "Task data source"; break} 
        "4" { "Powershell data source"; break}
    }
    
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/datasource/named/$DatasourceName")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
      
        if([string]::IsNullOrEmpty($response.dataSourceGUID)) {
            #Create DataSource
            $body = @{
                name               = $DatasourceName;
                type               = $DatasourceType;
                model              = (ConvertFrom-Json-WithEmptyArray($DatasourceModel));
                automationTaskGUID = $AutomationTaskGuid;
                value              = (ConvertFrom-Json-WithEmptyArray($DatasourceStaticValue));
                script             = $DatasourcePsScript;
                input              = (ConvertFrom-Json-WithEmptyArray($DatasourceInput));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
      
            $uri = ($script:PortalBaseUrl +"api/v1/datasource")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
              
            $datasourceGuid = $response.dataSourceGUID
            Write-Information "$datasourceTypeName '$DatasourceName' created$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        } else {
            #Get DatasourceGUID
            $datasourceGuid = $response.dataSourceGUID
            Write-Warning "$datasourceTypeName '$DatasourceName' already exists$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        }
    } catch {
      Write-Error "$datasourceTypeName '$DatasourceName', message: $_"
    }

    $returnObject.Value = $datasourceGuid
}

function Invoke-HelloIDDynamicForm {
    param(
        [parameter(Mandatory)][String]$FormName,
        [parameter(Mandatory)][String]$FormSchema,
        [parameter(Mandatory)][Ref]$returnObject
    )
    
    $FormName = $FormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/forms/$FormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }
    
        if(([string]::IsNullOrEmpty($response.dynamicFormGUID)) -or ($response.isUpdated -eq $true)) {
            #Create Dynamic form
            $body = @{
                Name       = $FormName;
                FormSchema = (ConvertFrom-Json-WithEmptyArray($FormSchema));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/forms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
    
            $formGuid = $response.dynamicFormGUID
            Write-Information "Dynamic form '$formName' created$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        } else {
            $formGuid = $response.dynamicFormGUID
            Write-Warning "Dynamic form '$FormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        }
    } catch {
        Write-Error "Dynamic form '$FormName', message: $_"
    }

    $returnObject.Value = $formGuid
}


function Invoke-HelloIDDelegatedForm {
    param(
        [parameter(Mandatory)][String]$DelegatedFormName,
        [parameter(Mandatory)][String]$DynamicFormGuid,
        [parameter()][Array][AllowEmptyString()]$AccessGroups,
        [parameter()][String][AllowEmptyString()]$Categories,
        [parameter(Mandatory)][String]$UseFaIcon,
        [parameter()][String][AllowEmptyString()]$FaIcon,
        [parameter()][String][AllowEmptyString()]$task,
        [parameter(Mandatory)][Ref]$returnObject
    )
    $delegatedFormCreated = $false
    $DelegatedFormName = $DelegatedFormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$DelegatedFormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }
    
        if([string]::IsNullOrEmpty($response.delegatedFormGUID)) {
            #Create DelegatedForm
            $body = @{
                name            = $DelegatedFormName;
                dynamicFormGUID = $DynamicFormGuid;
                isEnabled       = "True";
                useFaIcon       = $UseFaIcon;
                faIcon          = $FaIcon;
                task            = ConvertFrom-Json -inputObject $task;
            }
            if(-not[String]::IsNullOrEmpty($AccessGroups)) { 
                $body += @{
                    accessGroups    = (ConvertFrom-Json-WithEmptyArray($AccessGroups));
                }
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
    
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Information "Delegated form '$DelegatedFormName' created$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
            $delegatedFormCreated = $true

            $bodyCategories = $Categories
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$delegatedFormGuid/categories")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $bodyCategories
            Write-Information "Delegated form '$DelegatedFormName' updated with categories"
        } else {
            #Get delegatedFormGUID
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Warning "Delegated form '$DelegatedFormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
        }
    } catch {
        Write-Error "Delegated form '$DelegatedFormName', message: $_"
    }

    $returnObject.value.guid = $delegatedFormGuid
    $returnObject.value.created = $delegatedFormCreated
}


<# Begin: HelloID Global Variables #>
foreach ($item in $globalHelloIDVariables) {
	Invoke-HelloIDGlobalVariable -Name $item.name -Value $item.value -Secret $item.secret 
}
<# End: HelloID Global Variables #>


<# Begin: HelloID Data sources #>
<# Begin: DataSource "Exchange-user-generate-table-distributiongroups-manage-memberships" #>
$tmpPsScript = @'
<#----- Exchange On-Premises: Start -----#>
# Connect to Exchange
try{
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername,$adminSecurePassword
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
'@ 
$tmpModel = @'
[{"key":"name","type":0},{"key":"sAMAccountName","type":0}]
'@ 
$tmpInput = @'
[]
'@ 
$dataSourceGuid_1 = [PSCustomObject]@{} 
$dataSourceGuid_1_Name = @'
Exchange-user-generate-table-distributiongroups-manage-memberships
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_1_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_1) 
<# End: DataSource "Exchange-user-generate-table-distributiongroups-manage-memberships" #>

<# Begin: DataSource "Exchange-distributiongroup-generate-table-members" #>
$tmpPsScript = @'
<#----- Exchange On-Premises: Start -----#>
# Connect to Exchange
try{
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername,$adminSecurePassword
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck #-SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop 
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

# Read current distribution group
try{    
    
    $members = Get-DistributionGroupMember -Identity $datasource.selectedGroup.UserPrincipalName
    Write-Information -Message "Found distribution group [$($datasource.selectedGroup.displayName)]"
    
    $members = $members | Sort-Object -Property Displayname
    foreach($member in $members)
    {
        if($member.RecipientType -eq "UserMailbox")
        {
            $displayValue = $member.Displayname + " [" + $member.Samaccountname + "]"
            $returnObject = @{sAMAccountName=$member.Samaccountname;name=$displayValue;}
            Write-Output $returnObject
        }
    }    
    
} catch {
    Write-Information -Message "Could not find distribution group [$($datasource.selectedGroup.UserPrincipalName)]"
    Write-Error -Message "Error at line: $($_.InvocationInfo.ScriptLineNumber - 79): $($_.Exception.Message)"
    if($debug -eq $true){
        Write-Information -Message "$($_.Exception)"
    }
    Write-Information -Message "Failed to find distribution group [$($adUser.userPrincipalName)]"
    throw $_
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
'@ 
$tmpModel = @'
[{"key":"sAMAccountName","type":0},{"key":"name","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"selectedGroup","type":0,"options":1}]
'@ 
$dataSourceGuid_2 = [PSCustomObject]@{} 
$dataSourceGuid_2_Name = @'
Exchange-distributiongroup-generate-table-members
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_2_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_2) 
<# End: DataSource "Exchange-distributiongroup-generate-table-members" #>

<# Begin: DataSource "Exchange-distributiongroup-generate-table-wildcard" #>
$tmpPsScript = @'
<#----- Exchange On-Premises: Start -----#>
# Connect to Exchange
try{
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername,$adminSecurePassword
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
'@ 
$tmpModel = @'
[{"key":"displayName","type":0},{"key":"UserPrincipalName","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"searchValue","type":0,"options":1}]
'@ 
$dataSourceGuid_0 = [PSCustomObject]@{} 
$dataSourceGuid_0_Name = @'
Exchange-distributiongroup-generate-table-wildcard
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_0_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_0) 
<# End: DataSource "Exchange-distributiongroup-generate-table-wildcard" #>
<# End: HelloID Data sources #>

<# Begin: Dynamic Form "Exchange on-premise - Manage memberships distribution group " #>
$tmpSchema = @"
[{"label":"Select group","fields":[{"key":"searchfield","templateOptions":{"label":"Search","placeholder":""},"type":"input","summaryVisibility":"Hide element","requiresTemplateOptions":true,"requiresKey":true},{"key":"gridGroups","templateOptions":{"label":"Select group","required":true,"grid":{"columns":[{"headerName":"Display Name","field":"displayName"},{"headerName":"Mail address","field":"UserPrincipalName"}],"height":300,"rowSelection":"single"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_0","input":{"propertyInputs":[{"propertyName":"searchValue","otherFieldValue":{"otherFieldKey":"searchfield"}}]}},"useFilter":false},"type":"grid","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true}]},{"label":"Members","fields":[{"key":"members","templateOptions":{"label":"Manage group memberships","required":false,"filterable":true,"useDataSource":true,"dualList":{"options":[{"guid":"75ea2890-88f8-4851-b202-626123054e14","Name":"Apple"},{"guid":"0607270d-83e2-4574-9894-0b70011b663f","Name":"Pear"},{"guid":"1ef6fe01-3095-4614-a6db-7c8cd416ae3b","Name":"Orange"}],"optionKeyProperty":"sAMAccountName","optionDisplayProperty":"name","labelLeft":"Available","labelRight":"Member of"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_1","input":{"propertyInputs":[]}},"destinationDataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_2","input":{"propertyInputs":[{"propertyName":"selectedGroup","otherFieldValue":{"otherFieldKey":"gridGroups"}}]}},"useFilter":false},"type":"duallist","summaryVisibility":"Show","sourceDataSourceIdentifierSuffix":"source-datasource","destinationDataSourceIdentifierSuffix":"destination-datasource","requiresTemplateOptions":true,"requiresKey":true}]}]
"@ 

$dynamicFormGuid = [PSCustomObject]@{} 
$dynamicFormName = @'
Exchange on-premise - Manage memberships distribution group 
'@ 
Invoke-HelloIDDynamicForm -FormName $dynamicFormName -FormSchema $tmpSchema  -returnObject ([Ref]$dynamicFormGuid) 
<# END: Dynamic Form #>

<# Begin: Delegated Form Access Groups and Categories #>
$delegatedFormAccessGroupGuids = @()
if(-not[String]::IsNullOrEmpty($delegatedFormAccessGroupNames)){
    foreach($group in $delegatedFormAccessGroupNames) {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/groups/$group")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
            $delegatedFormAccessGroupGuid = $response.groupGuid
            $delegatedFormAccessGroupGuids += $delegatedFormAccessGroupGuid
            
            Write-Information "HelloID (access)group '$group' successfully found$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormAccessGroupGuid })"
        } catch {
            Write-Error "HelloID (access)group '$group', message: $_"
        }
    }
    if($null -ne $delegatedFormAccessGroupGuids){
        $delegatedFormAccessGroupGuids = ($delegatedFormAccessGroupGuids | Select-Object -Unique | ConvertTo-Json -Depth 100 -Compress)
    }
}

$delegatedFormCategoryGuids = @()
foreach($category in $delegatedFormCategories) {
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories/$category")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
        
        Write-Information "HelloID Delegated Form category '$category' successfully found$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    } catch {
        Write-Warning "HelloID Delegated Form category '$category' not found"
        $body = @{
            name = @{"en" = $category};
        }
        $body = ConvertTo-Json -InputObject $body -Depth 100

        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories")
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid

        Write-Information "HelloID Delegated Form category '$category' successfully created$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    }
}
$delegatedFormCategoryGuids = (ConvertTo-Json -InputObject $delegatedFormCategoryGuids -Depth 100 -Compress)
<# End: Delegated Form Access Groups and Categories #>

<# Begin: Delegated Form #>
$delegatedFormRef = [PSCustomObject]@{guid = $null; created = $null} 
$delegatedFormName = @'
Exchange on-premise - Manage memberships distribution group
'@
$tmpTask = @'
{"name":"Exchange on-premise - Manage memberships distribution group","script":"$VerbosePreference = \"SilentlyContinue\"\r\n$InformationPreference = \"Continue\"\r\n$WarningPreference = \"Continue\"\r\n\r\n# variables configured in form\r\n$groupName = $form.gridGroups.userPrincipalName\r\n$groupDisplayName = $form.gridGroups.displayName\r\n$usersToRemove = $form.members.rightToLeft\r\n$usersToAdd = $form.members.leftToRight\r\n\r\ntry {\r\n    \u003c#----- Exchange On-Premises: Start -----#\u003e\r\n    # Connect to Exchange\r\n    try {\r\n        $adminSecurePassword = ConvertTo-SecureString -String \"$ExchangeAdminPassword\" -AsPlainText -Force\r\n        $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername, $adminSecurePassword\r\n        $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck\r\n        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop \r\n        #-AllowRedirection\r\n        $session = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber\r\n\r\n        Write-Information \"Successfully connected to Exchange using the URI [$exchangeConnectionUri]\" \r\n    \r\n        $Log = @{\r\n            Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n            System            = \"Exchange On-Premise\" # optional (free format text) \r\n            Message           = \"Successfully connected to Exchange using the URI [$exchangeConnectionUri]\" # required (free format text) \r\n            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n            TargetDisplayName = $exchangeConnectionUri # optional (free format text) \r\n            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) \r\n        }\r\n        #send result back  \r\n        Write-Information -Tags \"Audit\" -MessageData $log\r\n    }\r\n    catch {\r\n        Write-Error \"Error connecting to Exchange using the URI [$exchangeConnectionUri]. Error: $($_.Exception.Message)\"\r\n        $Log = @{\r\n            Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n            System            = \"Exchange On-Premise\" # optional (free format text) \r\n            Message           = \"Failed to connect to Exchange using the URI [$exchangeConnectionUri].\" # required (free format text) \r\n            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n            TargetDisplayName = $exchangeConnectionUri # optional (free format text) \r\n            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) \r\n        }\r\n        #send result back  \r\n        Write-Information -Tags \"Audit\" -MessageData $log\r\n    }\r\n\r\n    if ($usersToAdd.count -gt 0) {\r\n        try {\r\n            Write-Information \"Starting to add distribution group [$groupName] to users [$($usersToAdd.sAMAccountName)]\"\r\n            #$usersToAddJson = $usersToAdd | ConvertFrom-Json        \r\n            foreach ($user in $usersToAdd) {\r\n                try {\r\n                    Add-DistributionGroupMember -Identity $groupName -Member $user.sAMAccountName -Confirm:$false -ErrorAction Stop\r\n                    Write-Information \"Finished adding $($user.sAMAccountName) to distribution group [$groupName]\"\r\n                    $Log = @{\r\n                        Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n                        System            = \"Exchange On-Premise\" # optional (free format text) \r\n                        Message           = \"Successfully added [$($user.sAMAccountName)] to distribution group [$groupName]\" # required (free format text) \r\n                        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                        TargetDisplayName = $($user.name) # optional (free format text) \r\n                        TargetIdentifier  = $groupDisplayName # optional (free format text) \r\n                    }\r\n                    #send result back  \r\n                    Write-Information -Tags \"Audit\" -MessageData $log       \r\n                }\r\n                catch {\r\n                    Write-Error \"Error adding $($user.name) to distribution group [$groupName]. Error: $($_.Exception.Message)\" \r\n                    $Log = @{\r\n                        Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n                        System            = \"Exchange On-Premise\" # optional (free format text) \r\n                        Message           = \"Failed to add [$($user.sAMAccountName)] to distribution group [$groupName]\" # required (free format text) \r\n                        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                        TargetDisplayName = $($user.name) # optional (free format text) \r\n                        TargetIdentifier  = $groupDisplayName # optional (free format text) \r\n                    }\r\n                    #send result back  \r\n                    Write-Information -Tags \"Audit\" -MessageData $log                    \r\n                }                         \r\n            }                        \r\n        }\r\n        catch {\r\n            Write-Error \"Could not add distribution group [$groupName] to users [$($usersToAdd.sAMAccountName)]. Error: $($_.Exception.Message)\"\r\n            $Log = @{\r\n                Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n                System            = \"Exchange On-Premise\" # optional (free format text) \r\n                Message           = \"Failed to add distribution group [$groupName] to users [$($usersToAdd.sAMAccountName)]\" # required (free format text) \r\n                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                TargetDisplayName = $($usersToAdd.name) # optional (free format text) \r\n                TargetIdentifier  = $groupDisplayName # optional (free format text) \r\n            }\r\n            #send result back  \r\n            Write-Information -Tags \"Audit\" -MessageData $log            \r\n        }\r\n    }\r\n\r\n\r\n    if ($usersToRemove.count -gt 0) {\r\n        try {\r\n            Write-Information \"Starting to remove distribution group [$groupName] from users [$($usersToRemove.sAMAccountName)]\"\r\n            #$usersToRemoveJson = $usersToRemove | ConvertFrom-Json            \r\n            foreach ($user in $usersToRemove) {\r\n                try {\r\n                    Remove-DistributionGroupMember -Identity $groupName -Member $user.sAMAccountName -Confirm:$false -ErrorAction Stop\r\n                    Write-Information \"Finished removing  [$($user.name)] from distribution group [$groupName]\" \r\n                    $Log = @{\r\n                        Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n                        System            = \"Exchange On-Premise\" # optional (free format text) \r\n                        Message           = \"Successfully removed [$($user.sAMAccountName)] from distribution group [$groupName]\" # required (free format text) \r\n                        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                        TargetDisplayName = $($user.name) # optional (free format text) \r\n                        TargetIdentifier  = $groupDisplayName # optional (free format text) \r\n                    }\r\n                    #send result back  \r\n                    Write-Information -Tags \"Audit\" -MessageData $log                    \r\n                }\r\n                catch {\r\n                    Write-Error \"Failed to remove  [$($user.sAMAccountName)] from distribution group [$groupName]. Error: $($_.Exception.Message)\" \r\n                    $Log = @{\r\n                        Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n                        System            = \"Exchange On-Premise\" # optional (free format text) \r\n                        Message           = \"Failed to remove  [$($user.sAMAccountName)] from distribution group [$groupName]\" # required (free format text) \r\n                        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                        TargetDisplayName = $($user.name) # optional (free format text) \r\n                        TargetIdentifier  = $groupDisplayName # optional (free format text) \r\n                    }\r\n                    #send result back  \r\n                    Write-Information -Tags \"Audit\" -MessageData $log \r\n                }\r\n            }\r\n        }\r\n        catch {\r\n            Write-Error \"Could not remove distribution group [$groupName] from users [$($usersToRemove.sAMAccountName)] Error: $($_.Exception.Message)\"            \r\n            $Log = @{\r\n                Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n                System            = \"Exchange On-Premise\" # optional (free format text) \r\n                Message           = \"Failed to remove distribution group [$groupName] from users [$($usersToRemove.sAMAccountName)]\" # required (free format text) \r\n                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                TargetDisplayName = $($usersToRemove.name) # optional (free format text) \r\n                TargetIdentifier  = $groupDisplayName # optional (free format text) \r\n            }\r\n            #send result back  \r\n            Write-Information -Tags \"Audit\" -MessageData $log            \r\n        }    \r\n    }\r\n}\r\ncatch {\r\n    Write-Error \"Could not set memberships on distribution group [$($groupName)]. Error: $($_.Exception.Message)\"    \r\n    $Log = @{\r\n        Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n        System            = \"Exchange On-Premise\" # optional (free format text) \r\n        Message           = \"Failed setting memberships on distribution group [$groupName].\" # required (free format text) \r\n        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $groupDisplayName # optional (free format text) \r\n        TargetIdentifier  = $groupName # optional (free format text) \r\n    }\r\n    #send result back  \r\n    Write-Information -Tags \"Audit\" -MessageData $log\r\n}\r\nfinally {\r\n    # Disconnect from Exchange\r\n    try {\r\n        Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop\r\n        Write-Information \"Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]\"     \r\n        $Log = @{\r\n            Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n            System            = \"Exchange On-Premise\" # optional (free format text) \r\n            Message           = \"Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]\" # required (free format text) \r\n            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n            TargetDisplayName = $exchangeConnectionUri # optional (free format text) \r\n            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) \r\n        }\r\n        #send result back  \r\n        Write-Information -Tags \"Audit\" -MessageData $log\r\n    }\r\n    catch {\r\n        Write-Error \"Error disconnecting from Exchange.  Error: $($_.Exception.Message)\"\r\n        $Log = @{\r\n            Action            = \"UpdateResource\" # optional. ENUM (undefined = default) \r\n            System            = \"Exchange On-Premise\" # optional (free format text) \r\n            Message           = \"Failed to disconnect from Exchange using the URI [$exchangeConnectionUri].\" # required (free format text) \r\n            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n            TargetDisplayName = $exchangeConnectionUri # optional (free format text) \r\n            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) \r\n        }\r\n        #send result back  \r\n        Write-Information -Tags \"Audit\" -MessageData $log \r\n    }\r\n    \u003c#----- Exchange On-Premises: End -----#\u003e\r\n}\r\n\r\n","runInCloud":false}
'@ 

Invoke-HelloIDDelegatedForm -DelegatedFormName $delegatedFormName -DynamicFormGuid $dynamicFormGuid -AccessGroups $delegatedFormAccessGroupGuids -Categories $delegatedFormCategoryGuids -UseFaIcon "True" -FaIcon "fa fa-users" -task $tmpTask -returnObject ([Ref]$delegatedFormRef) 
<# End: Delegated Form #>

