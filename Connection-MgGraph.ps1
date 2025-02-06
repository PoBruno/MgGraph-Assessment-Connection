
$tenantId = "tenantId"
$clientId = "clientId"
$certificateThumbprint = "certificateThumbprint"


# Gerando Token de Acesso
$CertificatePath = "cert:\currentuser\my\$CertificateThumbprint"
$Certificate = Get-Item $certificatePath
$Token = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -ClientCertificate $Certificate -ForceRefresh
$SecureToken = ConvertTo-SecureString $Token.AccessToken -AsPlainText -Force

# Connectando no Graph
Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $certificateThumbprint
Connect-MgGraph -AccessToken $SecureToken




