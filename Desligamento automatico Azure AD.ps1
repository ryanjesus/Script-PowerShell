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

    # Verificar se a data de desativação é hoje ou anterior a hoje
    if ($disableDate -le (Get-Date)) {
        # Desativar a conta de usuário
        $user = Get-AzureADUser -Filter "UserPrincipalName eq '$email'"
        if ($user) {
            $user | Set-AzureADUser -AccountEnabled $false
            Write-Host "Usuário $email desativado com sucesso."
        } else {
            Write-Host "Usuário $email não encontrado."
        }
    } else {
        Write-Host "Usuário $email será desativado em $disableDate."
    }
}

#Arquivo CSV
#Email,DisableDate
#usuario1@dominio.com,2024-05-23
#usuario2@dominio.com,2024-05-30
#usuario3@dominio.com,2024-06-10