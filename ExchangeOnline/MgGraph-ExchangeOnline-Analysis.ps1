<#
.SYNOPSIS
    Script com queries simples para análise de caixas de email no Exchange Online usando Microsoft Graph PowerShell.
    Autor: Bruno Gomes
    Repositório: https://github.com/pobruno/MgGraph-Assessment-Connection
.DESCRIPTION
    Este script contém exemplos práticos de queries para análise de caixas de email, incluindo:
    - Listagem de regras de pasta.
    - Filtragem de emails por data, assunto ou remetente.
    - Listagem de pastas e contagem de itens.
.NOTES
    Execute os blocos de código conforme necessário.
#>

# 1. Conectar ao Microsoft Graph (solicita permissões para leitura de e-mail e leitura de usuários)
# Connectando no Graph
Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $certificateThumbprint
Connect-MgGraph -AccessToken $SecureToken

# Verifique a conexão, por exemplo:
(Get-MgProfile).UserPrincipalName

# Defina o usuário-alvo (substitua pelo e-mail do usuário desejado)
$UserId = "first.last@domain.com"

# ---------------------------------------------------------------
# 2. Listar todas as pastas da caixa de correio do usuário
Write-Host "Listando pastas de caixa de correio para $($UserId):"
Get-MgUserMailFolder -UserId $UserId | Select-Object Id, DisplayName, TotalItemCount
# (Esta query mostra o ID, nome e a contagem de itens de cada pasta)
# Fonte: :contentReference[oaicite:0]{index=0}

# ---------------------------------------------------------------
# 3. Listar mensagens na pasta "Inbox" com filtro por data e remetente
# Exemplo: listar mensagens enviadas pelo próprio usuário entre 1 e 31 de janeiro de 2024
$startDate = "2024-01-01T00:00:00Z"
$endDate = "2024-01-31T23:59:59Z"
$filterQuery = "sender/emailAddress/address eq '$UserId' and sentDateTime ge $startDate and sentDateTime le $endDate"

Write-Host "Listando mensagens no Inbox enviadas pelo próprio usuário entre $startDate e $($endDate):"
Get-MgUserMailFolderMessage -UserId $UserId -MailFolderId "Inbox" -Filter $filterQuery |
Select-Object Subject, SentDateTime, @{Name = "Sender"; Expression = { $_.Sender.EmailAddress.Address } }
# (A query utiliza filtros OData para restringir os resultados)
# Fonte: :contentReference[oaicite:1]{index=1}

# ---------------------------------------------------------------
# 4. Listar regras (Inbox Rules) configuradas na pasta "Inbox"
Write-Host "Listando regras de Inbox para $($UserId):"
Get-MgUserMailFolderMessageRule -UserId $UserId -MailFolderId "Inbox" |
Select-Object Id, DisplayName, Sequence, IsEnabled
# (Essa query mostra as regras ativas e suas propriedades)
# Fonte: :contentReference[oaicite:2]{index=2}

# ---------------------------------------------------------------
# 5. Contar o número de mensagens em cada pasta
Write-Host "Contagem de mensagens por pasta:"
$folders = Get-MgUserMailFolder -UserId $UserId
foreach ($folder in $folders) {
    Write-Host "Pasta: $($folder.DisplayName) - Itens: $($folder.TotalItemCount)"
}
# (Esta parte itera por cada pasta e exibe a contagem de itens)

# ---------------------------------------------------------------
# 6. Listar mensagens com anexos na pasta "Inbox"
Write-Host "Listando mensagens com anexos na pasta Inbox:"
Get-MgUserMailFolderMessage -UserId $UserId -MailFolderId "Inbox" -Filter "hasAttachments eq true" |
Select-Object Subject, HasAttachments
# (Filtra mensagens que possuem anexos)

# ---------------------------------------------------------------
# 7. Consultar configurações de Litigation Hold do usuário
Write-Host "Verificando status de Litigation Hold para $($UserId):"
Get-MgUser -UserId $UserId -Property "LitigationHoldEnabled,LitigationHoldDuration,LitigationHoldDate" |
Select-Object UserPrincipalName, LitigationHoldEnabled, LitigationHoldDuration, LitigationHoldDate
# (Esta query mostra se o Litigation Hold está ativado e seus parâmetros)
# Fonte: :contentReference[oaicite:3]{index=3}

# ---------------------------------------------------------------
# Fim do Script
# Cada bloco pode ser executado individualmente para teste e análise gradual.
