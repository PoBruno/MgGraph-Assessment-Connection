<# 
Author: Bruno Gomes
.SYNOPSIS
Cria (ou atualiza) um app registration com as permissões necessárias para rodar o script de avaliação do tenant e faz upload de um certificado self-signed da máquina.

.DESCRIPTION
Conecta ao Azure AD e provisiona um app registration com as permissões adequadas. Se o app já existir (pelo nome configurado), ele apenas gera um novo certificado na máquina e o adiciona ao app, de forma que diferentes máquinas possam ter seus certificados registrados.
    
.Notes
GitHub: https://github.com/pobruno/MgGraph-Assessment-Connection
Linkedin: https://www.linkedin.com/in/brunopoleza/
#>

## Requer o módulo AzureAD
#Requires -Modules AzureAD

function New-AadApplicationCertificate {
    [CmdletBinding(DefaultParameterSetName = 'DefaultSet')]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$CertificatePassword,

        [Parameter(Mandatory = $true, ParameterSetName = 'ClientIdSet')]
        [string]$ClientId,

        [string]$CertificateName,

        [Parameter(Mandatory = $false, ParameterSetName = 'ClientIdSet')]
        [switch]$AddToApplication
    )
    # Cria um certificado self-signed com validade de 2 anos
    $notAfter = (Get-Date).AddYears(2)
    try {
        $cert = New-SelfSignedCertificate -DnsName "MgGraph-Tool" -CertStoreLocation "cert:\CurrentUser\My" -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter $notAfter
    }
    catch {
        Write-Error "ERRO. Talvez seja necessário executar como Administrador."
        Write-Host $_
        return
    }

    if ($AddToApplication) {
        # Converte o certificado para Base64 e o adiciona ao app registration
        $KeyValue = [System.Convert]::ToBase64String($cert.GetRawCertData())
        New-AzureADApplicationKeyCredential -ObjectId $appReg.ObjectId -Type AsymmetricX509Cert -Usage Verify -Value $KeyValue | Out-Null
    }
    return $cert.Thumbprint
}

## Declaração de variáveis
$connected = $false
$appName = "MgGraph-Tool"
$appURI = @("https://localhost")
$appReg = $null
$ConsentURl = "https://login.microsoftonline.com/{tenant-id}/adminconsent?client_id={client-id}"
$TenantID = $null

## Tenta conectar ao Azure AD até ser bem-sucedido
while (-not $connected) {
    try {
        Connect-AzureAD -ErrorAction Stop
        $connected = $true
    }
    catch {
        Write-Host "Erro ao conectar ao Azure AD:`n$($error[0])`nTente novamente..." -ForegroundColor Red
        $connected = $false
    }
}

## Prepara o objeto de permissão (RequiredResourceAccess)
try {
    $Permissions = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
    # Lista de permissões (IDs dos recursos) conforme referência em https://docs.microsoft.com/en-us/graph/permissions-reference
    $permList = @(
        "332a536c-c7ef-4017-ab91-336970924f0d",
        "246dd0d5-5bd0-4def-940b-0421030a5b68",
        "01d4889c-1287-42c6-ac1f-5d1e02578ef6",
        "5b567255-7703-4780-807c-7be8301ae99b",
        "7ab1d382-f21e-4acd-a863-ba3e13f7da61",
        "2280dda6-0bfd-44ee-a2f4-cb867cfc4c1e",
        "230c1aed-a721-4c5d-9cb4-a90514e508ef",
        "37730810-e9ba-4e46-b07e-8ca78d182097",
        "59a6b24b-4225-4393-8165-ebaec5f55d7a"
    )
    $permArray = @()
    foreach ($perm in $permList) {
        $permArray += New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList $perm, "Role"
    }
    # Obter o Service Principal do Exchange Online (AppID fixo)
    $EXOapi = Get-AzureADServicePrincipal -Filter "AppID eq '00000002-0000-0ff1-ce00-000000000000'"
    $EXOpermission = $EXOapi.AppRoles | Where-Object { $_.Value -eq 'Exchange.ManageAsApp' }
    $EXOapiPermission = [Microsoft.Open.AzureAD.Model.RequiredResourceAccess]@{
        ResourceAppId  = $EXOapi.AppId;
        ResourceAccess = [Microsoft.Open.AzureAD.Model.ResourceAccess]@{
            Id   = $EXOpermission.Id;
            Type = "Role"
        }
    }
    $permissions.ResourceAccess = $permArray
    $permissions.ResourceAppId = "00000003-0000-0000-c000-000000000000"
}
catch {
    Write-Host "Erro ao preparar o script:`n$($error[0])`nVerifique os pré-requisitos`nSaindo..." -ForegroundColor Red
    pause
    exit
}

## Verifica se já existe um app registration com o mesmo nome
$appReg = Get-AzureADApplication -Filter "DisplayName eq '$appName'" -ErrorAction SilentlyContinue

if ($appReg) {
    Write-Host "App registration '$appName' ja existe. Atualizando com novo certificado..." -ForegroundColor Yellow
}
else {
    try {
        # Cria novo app registration com as permissões necessárias
        $appReg = New-AzureADApplication -DisplayName $appName -ReplyUrls $appURI -ErrorAction Stop -RequiredResourceAccess $Permissions, $EXOapiPermission
        Write-Host "Aguardando o provisionamento do app..."
        Start-Sleep -Seconds 20
        # Habilita o Service Principal
        $SP = New-AzureADServicePrincipal -AppID $appReg.AppID
        ## Adiciona a role Global Reader
        $directoryRole = 'Global Reader'
        $RoleId = (Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -eq $directoryRole }).ObjectID
        if (-not $RoleId) {
            Write-Host "Funcao Global Reader ainda nao provisionada - Provisionando"
            $template = Get-AzureADDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq $directoryRole }
            Enable-AzureADDirectoryRole -RoleTemplateId $template.ObjectId
            $RoleId = (Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -eq $directoryRole }).ObjectID
        }
        Add-AzureADDirectoryRoleMember -ObjectId $RoleId -RefObjectId $SP.ObjectID -Verbose

        ## Adiciona a role Exchange Administrator
        $directoryRole = 'Exchange Administrator'
        $RoleId = (Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -eq $directoryRole }).ObjectID
        if (-not $RoleId) {
            Write-Host "Funcao Exchange Administrator ainda não provisionada - Provisionando"
            $template = Get-AzureADDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq $directoryRole }
            Enable-AzureADDirectoryRole -RoleTemplateId $template.ObjectId
            $RoleId = (Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -eq $directoryRole }).ObjectID
        }
        Add-AzureADDirectoryRoleMember -ObjectId $RoleId -RefObjectId $SP.ObjectID -Verbose

        ## Obtém o Tenant ID
        $TenantID = (Get-AzureADTenantDetail).ObjectId

        ## Atualiza a URL de consentimento
        $ConsentURl = $ConsentURl.Replace('{tenant-id}', $TenantID)
        $ConsentURl = $ConsentURl.Replace('{client-id}', $appReg.AppId)

        Write-Host "A pagina de consentimento sera aberta. Nao se esqueca de logar como admin para conceder o consentimento!" -ForegroundColor Yellow
        Start-Process $ConsentURl
    }
    catch {
        Write-Host "Erro ao criar novo app registration:`n$($error[0])`nSaindo..." -ForegroundColor Red
        pause
        exit
    }
}

## Cria um novo certificado na máquina e faz upload para o app
$Thumbprint = New-AadApplicationCertificate -ClientId $appReg.AppId -CertificatePassword 'Pa$$w0rd' -AddToApplication -CertificateName "Tenant Assessment Certificate"
$TenantID = (Get-AzureADTenantDetail).ObjectId

Write-Host "`nOs detalhes abaixo podem ser usados para conexao com MgGraph.:`n`n`tTenant ID:`t $TenantID`n`tClient ID:`t $($appReg.AppId)`n`tCertThumbprint:`t $Thumbprint`n" -ForegroundColor Green
Pause
