param (
    [string]$subject,
    [string]$recipients,
    [string]$body,
    [string]$filePath
)

Write-Host "DEBUG: Script reached PowerShell"
Write-Host "Subject: $subject"
Write-Host "Recipients: $recipients"
Write-Host "Body: $body"
Write-Host "File Path: $filePath"

$recipients.Split(",") | ForEach-Object {
    Write-Host "DEBUG: Recipient: $_"
}

Write-Host "DEBUG: Script finished successfully"
