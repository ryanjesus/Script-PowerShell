# Instale o módulo Azure Active Directory se você ainda não o tiver instalado
Install-Module -Name AzureAD

# Conecte-se ao Azure AD
Connect-AzureAD

# Caminho para o arquivo CSV contendo os endereços de e-mail e datas de desativação
$csvPath = "Caminho\para\seu\arquivo.csv"

# Carregar os dados do arquivo CSV
$data = Import-Csv -Path $csvPath

# Iterar sobre cada linha do CSV
foreach ($entry in $data) {
    $email = $entry.Email
    $disableDate = Get-Date $entry.DisableDate

    # Desativar usuário no Active Directory local por data
    if ($disableDate -le (Get-Date)) {
        Disable-ADAccount -Identity $email
        Write-Host "Usuário $email desativado com sucesso no Active Directory local."
    } else {
        Write-Host "Usuário $email será desativado em $disableDate no Active Directory local."
    }

    # Desativar usuário no Azure Active Directory
    $userAzure = Get-AzureADUser -Filter "UserPrincipalName eq '$email'"
    if ($userAzure) {
        $userAzure | Set-AzureADUser -AccountEnabled $false
        Write-Host "Usuário $email desativado com sucesso no Azure AD."
    } else {
        Write-Host "Usuário $email não encontrado no Azure AD."
    }
}