$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form
$username = $form.gridmailuser.userPrincipalName
$devicesToDelete = $form.selectedDevices

try {
    <#----- Exchange On-Premises: Start -----#>
    # Connect to Exchange
    try {
        $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
        $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername, $adminSecurePassword
        $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop 
        #-AllowRedirection
        $session = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber

        Write-Information "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" 
    
        $Log = @{
            Action            = "DeleteResource" # optional. ENUM (undefined = default) 
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

    if ($devicesToDelete.count -gt 0) {
        try {
            Write-Information "Starting to delete device(s) [$($devicesToDelete.FriendlyName)] for user [$username)]"
            
            foreach ($device in $devicesToDelete) {
                try {
                    Remove-ActiveSyncDevice -Identity $($device.DeviceI) -Confirm:$false

                    Write-Information "Finished deleting $($device.DeviceId) for user [$username]"
                    $Log = @{
                        Action            = "DeleteResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Successfully deleted $($device.DeviceId) for user [$username]" # required (free format text) 
                        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $username # optional (free format text) 
                        TargetIdentifier  = $($device.DeviceId) # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log       
                }
                catch {
                    Write-Error "Error deleting $($device.DeviceId) for user [$username]. Error: $($_.Exception.Message)" 
                    $Log = @{
                        Action            = "DeleteResource" # optional. ENUM (undefined = default) 
                        System            = "Exchange On-Premise" # optional (free format text) 
                        Message           = "Failed to allow [$($device.DeviceId)] for [$username]" # required (free format text) 
                        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                        TargetDisplayName = $username # optional (free format text) 
                        TargetIdentifier  = $($device.DeviceId) # optional (free format text) 
                    }
                    #send result back  
                    Write-Information -Tags "Audit" -MessageData $log                    
                }
            }
        }               
        catch {
            Write-Error "Could not delete device(s) [$($devicesToDelete.FriendlyName)] for user [$username]. Error: $($_.Exception.Message)"
            $Log = @{
                Action            = "DeleteResource" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premise" # optional (free format text) 
                Message           = "Failed to allow [$($devicesToDelete.FriendlyName)] for user [$username]" # required (free format text) 
                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $username # optional (free format text) 
                TargetIdentifier  = $($devicesToDelete.DeviceId) # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log            
        }
    } 
} catch {
    Write-Error "Could not delete devices for user [$username]. Error: $($_.Exception.Message)"    
    $Log = @{
        Action            = "DeleteResource" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Failed to delete devices for user [$username]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $username # optional (free format text) 
        TargetIdentifier  = $username # optional (free format text) 
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
            Action            = "DeleteResource" # optional. ENUM (undefined = default) 
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


