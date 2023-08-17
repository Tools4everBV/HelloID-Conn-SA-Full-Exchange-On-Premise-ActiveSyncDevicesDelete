<#----- Exchange On-Premises: [powershell-datasource]_Exchange-mailbox-add-email-address-get-mailbox -----#>
# Connect to Exchange
try {
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername, $adminSecurePassword
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck #-SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos -ErrorAction Stop
    Write-Information "Successfully connected to Exchange using the URI [$exchangeConnectionUri]"
} catch {
    Write-Information "Error connecting to Exchange using the URI [$exchangeConnectionUri]"
    Write-Information "Failed to connect to Exchange using the URI [$exchangeConnectionUri]"
    Write-Error "$($_.Exception.Message)"
    throw $_
}

try {
    $ParamsGetMailbxox = @{
       Mailbox = $dataSource.selecteduser.UserPrincipalName
    }
    
    Write-Information "SearchQuery: $($ParamsGetMailbxox.Filter)"
        $devices = Invoke-Command -Session $exchangeSession -ScriptBlock {
            Param ($ParamsGetMailbxox)
            Get-MobileDevice @ParamsGetMailbxox | Select-Object FriendlyName,UserDisplayName,DeviceFriendlyName,DevicePhoneNumber,DeviceId,DeviceOS,DeviceType,DeviceUserAgent,DeviceModel,DistinguishedName,FirstSyncTime,DeviceAccessState,DeviceAccessStateReason,DeviceAccessControlRule,ClientType
        } -ArgumentList $ParamsGetMailbxox

    
        $devices = $devices | Sort-Object -Property FriendlyName
        $resultCount = @($devices).Count
        Write-Information "Result count: $resultCount"
        Write-Information ($devices[0] | ConvertTo-Json)
        if ($resultCount -gt 0) {
            foreach ($device in $devices) {
                
                    $returnObject = @{DeviceId = $device.DeviceId; FriendlyName = $device.FriendlyName; DeviceType = $device.DeviceType; DeviceModel=$device.DeviceModel; State=$device.DeviceAccessState.Value }
                    Write-Output $returnObject
                
            }
        }
    
} catch {
    Write-Error "Error searching AD user [$searchValue]. Error: $($_.Exception.Message)"
}

# Disconnect from Exchange
try {
    Remove-PSSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
    Write-Information "Successfully disconnected from Exchange"
} catch {
    Write-Error "Error disconnecting from Exchange"
    Write-Error "$($_.Exception.Message)"
    throw $_
}
<#----- Exchange On-Premises: End -----#>

