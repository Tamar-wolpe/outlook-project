param(
    [string]$subject,
    [string]$recipients,
    [string]$body,
    [string]$filePath
)

try {
    # בדיקה איזו גרסת Outlook מותקנת
    $outlookPath = if (Test-Path "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE") {
        "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
    } elseif (Test-Path "C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE") {
        "C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE"
    } else {
        "OUTLOOK.EXE"  # שימוש ב-PATH
    }

    # הפעלת Outlook אם לא רץ
    if (-not (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue)) {
        Start-Process $outlookPath
        Start-Sleep -Seconds 3  # מחכים ש-Outlook יטען
    }

    # חיבור ל-Outlook
    $outlook = New-Object -ComObject Outlook.Application
    
    # יצירת הודעת דואר לכל נמען
    $recipientsList = $recipients.Split(",") | ForEach-Object { $_.Trim() }
    foreach ($r in $recipientsList) {
        $mail = $outlook.CreateItem(0)  # 0 = olMailItem
        $mail.Subject = $subject
        $mail.Body = $body
        $mail.To = $r
        
        if ($filePath -and (Test-Path $filePath)) {
            $mail.Attachments.Add($filePath) | Out-Null
        }
        
        $mail.Display($false)  # הצגת ההודעה
    }

    Write-Output "SUCCESS: טיוטות המייל נפתחו בהצלחה"
    
} catch {
    Write-Error "ERROR: $($_.Exception.Message)"
    exit 1
}