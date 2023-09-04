## connect to exchange and get list of mailboxes
try{
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername,$adminSecurePassword)
    $searchValue = ($dataSource.searchmailuser).trim()
    $searchQuery = "*$searchValue*"  

    $sessionOptionParams = @{
        SkipCACheck = $true
        SkipCNCheck = $true
        SkipRevocationCheck = $true
    }

    $sessionOption = New-PSSessionOption  @SessionOptionParams 

    $sessionParams = @{        
        Authentication = 'Kerberos' 
        ConfigurationName = 'Microsoft.Exchange' 
        ConnectionUri = $ExchangeConnectionUri 
        Credential = $adminCredential        
        SessionOption = $sessionOption       
    }

    $exchangeSession = New-PSSession @SessionParams

    Write-Information "Search query is '$searchQuery'" 
    
    $getMailboxParams = @{
        #RecipientTypeDetails = @('UserMailbox')         
        Filter = "Alias -like '$searchQuery' -or Name -like '$searchQuery'"   
    }
   
    
     $invokecommandParams = @{
        Session = $exchangeSession
        Scriptblock = [scriptblock] { Param ($Params)Get-Mailbox @Params}
        ArgumentList = $getMailboxParams
    }

    Write-Information "Successfully connected to Exchange '$ExchangeConnectionUri'"  
    
    $mailBoxes =  Invoke-Command @invokeCommandParams   

    $resultMailboxList = [System.Collections.Generic.List[PSCustomObject]]::New()
    foreach ($box in $mailBoxes)
    {        
       $resultMailbox = @{
        DisplayName = $box.DisplayName        
        UserPrincipalName = $box.UserPrincipalName
        Alias = $box.Alias
        DistinguishedName = $box.DistinguishedName        

       }
       $resultMailboxList.add($resultMailbox)

    }
    $resultMailboxList
    
    Remove-PSSession($exchangeSession)
  
} catch {
    Write-Error "Error connecting to Exchange using the URI '$exchangeConnectionUri', Message '$($_.Exception.Message)'"
}

