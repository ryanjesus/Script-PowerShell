Set-ExecutionPolicy RemoteSigned
Install-Module ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement

Connect-ExchangeOnline

Start-Process powershell ise -Verb RunAs #Power Shell com administrador

Remove-RecipientPermission "UPN do grupo" -AccessRights SendAs -Trustee user@contoso.com #Remover colaborador do grupo

Get-RecipientPermission "user@contoso.com" # Pega permissão

Add-RecipientPermission "UPN do grupo" -AccessRights SendAs -Trustee user@contoso.com 
 
Remove-RecipientPermission "UPN do grupo" -AccessRights SendAs -Trustee user@contoso.com
