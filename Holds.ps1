#TODO por hacer

# Conectarse al servicio de Exchange Online de Microsoft 365
Connect-ExchangeOnline -UserPrincipalName admin@tudominio.com

# Definir el correo electrónico de la cuenta que deseas quitar el hold
$CorreoElectronico = "usuario@tudominio.com"

# Buscar y obtener información de la cuenta
$Cuenta = Get-Mailbox -Identity $CorreoElectronico

# Desactivar todos los tipos de hold de la cuenta
Set-Mailbox -Identity $Cuenta.Identity -LitigationHoldEnabled $false -InPlaceHoldsEnabled $false -RetentionHoldEnabled $false -LitigationHoldDate $null -RetentionHoldDate $null
