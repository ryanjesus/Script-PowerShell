Install-Module -Name ImportExcel
Connect-ExchangeOnline

#Variavel de email
$email = "user@contoso.com"
$inicio = Get-Date -Date "2024-05-10" -Hour 00 -Minute 00 -Second 00
$fim = Get-Date -Date "2024-05-10" -Hour 23 -Minute 59 -Second 00
$mensagem = Get-MessageTrace -StartDate $inicio -EndDate $fim -RecipientAddress $email
$qtd = $mensagem.Count

#Abrir o power shell como administrador
Start-Process powershell ise -Verb RunAs

#Mostra os emails
#Get-MessageTrace -StartDate "2024-05-08" -Hour 00 -Minute 00 -Second 00 -RecipientAddress $email

#Get-MessageTrace -EndDate "2024-05-09" -Hour 00 -Minute 00 -Second 00 -RecipientAddress $email


#Exibição na tela
Write-Host "O endereço $email recebeu $qtd Verificação iniciado $inicio Verifica finalizado $fim"
Get-MessageTrace -StartDate $inicio -EndDate $fim -RecipientAddress $email

#Criar objeto Excel
$excelOutput = @()

# Loop através de cada mensagem e extrari as informações relevantes
foreach ($messages in $mensagem) {
    $data = [PSCustomObject]@{
        'Data e Hora' = $messages.Received
        'Remetente' = $messages.SenderAddress
        'Assunto' = $messages.Subject
        'Tamnho' = $messages.Size
        'Prioridade' = $messages.Priority
    }
    $excelOutput += $data
}

#Exportar para Excel
$excelOutput | Export-Excel -Path "C:\temp\Relatorio.xlsx" -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName "Relatório de Emails"

# Endereço de e-mail do remetente
$senderEmail = "user@contoso.com"

# Endereço de e-mail do destinatário
$recipientEmail = "user@contoso.com"

# Assunto do e-mail
$subject = "Relatório de E-mails"

# Corpo do e-mail
$body = "Por favor, encontre em anexo o relatório de e-mails."

# Caminho para o arquivo Excel
$attachmentPath = "C:\temp\Relatorio.xlsx"

# Configurações do servidor SMTP
$smtpServer = "smtp.dominio.com"
$smtpPort = 587  # Porta SMTP (normalmente 587 para TLS)

# Enviar o e-mail
Send-MailMessage -From $senderEmail -To $recipientEmail -Subject $subject -Body $body -SmtpServer $smtpServer -Port $smtpPort -Attachments $attachmentPath

