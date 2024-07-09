@{
    <#
        Custom-Alerts.psd1
        - Version: 1.0
        - Last Modified: 07/09/2024
        - SECOUGHL : File creation and initial alert additions        
    #>
    Alert = @(
      @{ 
        Path = '/Accounting/Accounts Payable'
        Email = 'john.doe@contoso.com','payable_dl@contoso.com'
       }
      @{ 
        Path = 'Accounting/Accounts Receivable'
        Email = 'jane.doe@contoso.com','receivable_dl@contoso.com'
       }  
        
    )
}
