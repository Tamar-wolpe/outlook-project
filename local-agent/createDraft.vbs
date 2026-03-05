' CreateDraft.vbs
' This script creates an Outlook draft with the specified parameters

' Enable error handling
On Error Resume Next

' Get command line arguments
Set objArgs = WScript.Arguments
If objArgs.Count < 4 Then
    WScript.Echo "ERROR: Missing required arguments"
    WScript.Quit 1
End If

' Create filesystem object first
Set fso = CreateObject("Scripting.FileSystemObject")

' Create log directory if it doesn't exist
logDir = "C:\temp"
If Not fso.FolderExists(logDir) Then
    fso.CreateFolder(logDir)
End If

' Read parameters - use WScript.Arguments.Unnamed to handle spaces in paths
subject = objArgs(0)
recipients = objArgs(1)
body = objArgs(2)
filePath = objArgs(3)

' Log the parameters (for debugging)
Set logFile = fso.OpenTextFile(logDir & "\outlook_script.log", 8, True) ' 8 = ForAppending
logFile.WriteLine("[" & Now() & "] Starting script with parameters:")
logFile.WriteLine("  Subject: " & subject)
logFile.WriteLine("  Recipients: " & recipients)
logFile.WriteLine("  Body length: " & Len(body))
logFile.WriteLine("  File path: " & filePath)

' Create Outlook application object
Set outlook = CreateObject("Outlook.Application")
If Err.Number <> 0 Then
    logFile.WriteLine("ERROR: Could not create Outlook object: " & Err.Description)
    WScript.Echo "ERROR: Could not create Outlook object: " & Err.Description
    WScript.Quit 1
End If

' Process each recipient
recipientList = Split(recipients, ",")
For Each recipient In recipientList
    recipient = Trim(recipient)
    If recipient <> "" Then
        logFile.WriteLine("Processing recipient: " & recipient)
        
        ' Create a new mail item
        Set mail = outlook.CreateItem(0) ' olMailItem = 0
        
        ' Set mail properties
        mail.Subject = subject
        mail.Body = body
        mail.To = recipient
        
        ' Add attachment if specified
        If filePath <> "" And filePath <> "undefined" Then
            logFile.WriteLine("  Checking file: " & filePath)
            
            ' Check if file exists (handle paths with spaces)
            If fso.FileExists(filePath) Then
                logFile.WriteLine("  Adding attachment: " & filePath)
                
                ' Add the attachment
                On Error Resume Next
                Set attachment = mail.Attachments.Add(filePath)
                If Err.Number <> 0 Then
                    logFile.WriteLine("  ERROR adding attachment: " & Err.Description)
                    WScript.Echo "ERROR: " & Err.Description
                    WScript.Quit 1
                End If
                On Error Goto 0
            Else
                logFile.WriteLine("  WARNING: File not found: " & filePath)
            End If
        End If
        
        ' Display the email (as a draft)
        mail.Display False ' False = don't show the inspector window
        
        logFile.WriteLine("  Draft created for: " & recipient)
    End If
Next

logFile.WriteLine("Script completed successfully")
logFile.Close

WScript.Echo "SUCCESS: Drafts created successfully"
WScript.Quit 0