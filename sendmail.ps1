param(
    [string]$To,
    [string]$Subject,
    [string]$Body
)

try {
    $app = New-Object -ComObject Outlook.Application
    $mail = $app.CreateItem(0)  # 0 = olMailItem
    $mail.Subject = $Subject
    $mail.Body = $Body
    $mail.To = $To
    $mail.Send()
    Write-Host "Mail sent successfully"
}
catch {
    Write-Error "Failed to send mail: $_"
    exit 1
}
