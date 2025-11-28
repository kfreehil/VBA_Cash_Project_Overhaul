'C:\Windows\System32\WScript.exe = WScript.exe
Dim ScriptHost
ScriptHost = Mid(WScript.FullName, InStrRev(WScript.FullName, "\") + 1, Len(WScript.FullName))

Dim oWs : Set oWs = CreateObject("WScript.Shell")
Dim oProcEnv : Set oProcEnv = oWs.Environment("Process")

' Am I running 64-bit version of WScript.exe/Cscript.exe? So, call #cript again in x86 script host and then exit.
If InStr(LCase(WScript.FullName), LCase(oProcEnv("windir") & "\System32\")) And oProcEnv("PROCESSOR_ARCHITECTURE") = "AMD64" Then
    'rebuild arguments
    If Not WScript.Arguments.Count = 0 Then
        Dim sArg, Arg
        sArg = ""
        For Each Arg In Wscript.Arguments
            'msgbox Arg,64
            sArg = sArg & " " & """" & Arg & """"
        Next
    End If

    ' rewriting command
    Dim sCmd : sCmd = """" & oProcEnv("windir") & "\SysWOW64\" & ScriptHost & """" & " """ & WScript.ScriptFullName & """" & sArg
    
    'msgbox "Call " & sCmd, 64
    'WScript.Echo "Call " & sCmd
    oWs.Run sCmd
    WScript.Quit
End If

'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

Dim StartTime
Dim Elapsed

StartTime = Timer

Dim DATA_DIR_WOODSON
'DATA_DIR_WOODSON = "\\pc.internal.macquarie.com\FSVC\AMERICAS\wnycfsp28535\" & _
'                    "Shared\tnc\SCF\Commodity Margin Lending\CML Start-up\" & _
'                    "Middle Office Procedures\Woodson\VBA Examples\data\"

DATA_DIR_WOODSON = "C:\Users\kfreehil\MyData\"
Const DB_BONY = "BONYNostro.accdb"

Dim AccessDB
Set AccessDB = CreateObject("Access.Application")
With AccessDB
    .OpenCurrentDatabase DATA_DIR_WOODSON & DB_BONY
    .Run "IngestNewData", False
    .Quit
End With
Set AccessDB = Nothing

Dim AccessDB2
Set AccessDB2 = CreateObject("Access.Application")
Const DB_MANAGER = "Manager.accdb"
With AccessDB2
    .OpenCurrentDatabase DATA_DIR_WOODSON & DB_MANAGER
    .Run "ConfirmtBony"
    .Quit
End With
Set AccessDB2 = Nothing

'Dim ProgramPath, WshShell, ProgramArgs, WaitOnReturn,intWindowStyle
'Set WshShell=CreateObject ("WScript.Shell")
'ProgramPath="C:\Users\kfreehil\MyData\ConfirmBONYSettlements.vbs"
'ProgramArgs=""
'intWindowStyle=1
'WaitOnReturn=True
'WshShell.Run Chr (34) & ProgramPath & Chr (34) & Space (1) & ProgramArgs,intWindowStyle, WaitOnReturn
'Set wshShell = Nothing

'═══════════════════════════════════════════════════════════════════════════════════════════════════
' DAILY MAINTENANCE (runs ONCE per day, BEFORE opening database)
' Compact happens while database is CLOSED - no permission errors!
'═══════════════════════════════════════════════════════════════════════════════════════════════════
If IsDailyMaintenanceNeeded(DATA_DIR_WOODSON) Then
    PerformDailyMaintenance dbPath, DATA_DIR_WOODSON
End If


Elapsed = Timer - StartTime

msgbox "BONY Data has been loaded via Outlook! Timer: " & PrintHrMinSec(Elapsed),64

'═══════════════════════════════════════════════════════════════════════════════════════════════════
'═══════════════════════════════════════════════════════════════════════════════════════════════════
' MAINTENANCE FUNCTIONS
'═══════════════════════════════════════════════════════════════════════════════════════════════════
'═══════════════════════════════════════════════════════════════════════════════════════════════════

'**********************
'*** CHECK IF MAINTENANCE NEEDED ***
'**********************
Function IsDailyMaintenanceNeeded(dataDir)
    ' Check if maintenance has run today by reading a tracking file
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim trackingFile
    trackingFile = dataDir & "LastMaintenanceDate.txt"
    
    ' File doesn't exist - maintenance needed
    If Not fso.FileExists(trackingFile) Then
        IsDailyMaintenanceNeeded = True
        Set fso = Nothing
        Exit Function
    End If
    
    ' Read last maintenance date
    Dim ts, lastDateStr
    Set ts = fso.OpenTextFile(trackingFile, 1) ' 1 = ForReading
    lastDateStr = Trim(ts.ReadLine)
    ts.Close
    
    ' Compare to today (date only, not time)
    Dim lastDate
    lastDate = CDate(lastDateStr)
    
    If DateValue(lastDate) < DateValue(Now) Then
        IsDailyMaintenanceNeeded = True
    Else
        IsDailyMaintenanceNeeded = False
    End If
    
    Set fso = Nothing
End Function

'**********************
'*** PERFORM DAILY MAINTENANCE ***
'**********************
Sub PerformDailyMaintenance(dbPath, dataDir)
    WScript.Echo "==========================================="
    WScript.Echo "       DAILY MAINTENANCE"
    WScript.Echo "==========================================="
    WScript.Echo ""
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim maintenanceStart
    maintenanceStart = Timer
    
    '───────────────────────────────────────────
    ' STEP 1: Compact Database (while CLOSED)
    '───────────────────────────────────────────
    WScript.Echo "Step 1: Compacting database..."
    
    Dim tempPath, backupPath
    tempPath = Replace(dbPath, ".accdb", "_temp.accdb")
    backupPath = Replace(dbPath, ".accdb", "_backup.accdb")
    
    ' Get size before
    Dim sizeBefore
    sizeBefore = fso.GetFile(dbPath).Size
    WScript.Echo "  Size before: " & FormatNumber(sizeBefore / 1024 / 1024, 1) & " MB"
    
    ' Compact using DAO (database must be CLOSED)
    Dim engine
    Set engine = CreateObject("DAO.DBEngine.120")
    
    Dim compactSuccess
    compactSuccess = False
    
    On Error Resume Next
    
    ' Delete temp file if exists from previous failed attempt
    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath
    
    ' Compact to temp file
    engine.CompactDatabase dbPath, tempPath
    
    If Err.Number = 0 Then
        ' Success - swap files
        If fso.FileExists(backupPath) Then fso.DeleteFile backupPath
        fso.MoveFile dbPath, backupPath
        fso.MoveFile tempPath, dbPath
        
        Dim sizeAfter
        sizeAfter = fso.GetFile(dbPath).Size
        WScript.Echo "  Size after: " & FormatNumber(sizeAfter / 1024 / 1024, 1) & " MB"
        WScript.Echo "  Saved: " & FormatNumber((sizeBefore - sizeAfter) / 1024 / 1024, 1) & " MB"
        WScript.Echo "  Backup: " & backupPath
        WScript.Echo "  [OK] Compact complete!"
        compactSuccess = True
    Else
        WScript.Echo "  [ERROR] Compact failed: " & Err.Description
        Err.Clear
        ' Clean up temp file if exists
        If fso.FileExists(tempPath) Then fso.DeleteFile tempPath
    End If
    
    On Error GoTo 0
    Set engine = Nothing
    WScript.Echo ""
    
    '───────────────────────────────────────────
    ' STEP 2: Refresh Index (requires opening database briefly)
    '───────────────────────────────────────────
    WScript.Echo "Step 2: Refreshing ValueDate index..."
    
    Dim indexStart
    indexStart = Timer
    
    Dim accessApp
    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase dbPath
    
    ' Drop existing index (ignore error if doesn't exist)
    On Error Resume Next
    accessApp.CurrentDb.Execute "DROP INDEX idx_valuedate ON BonyStatement"
    Err.Clear
    On Error GoTo 0
    
    ' Create fresh index
    accessApp.CurrentDb.Execute "CREATE INDEX idx_valuedate ON BonyStatement(ValueDate)"
    
    ' Close Access
    accessApp.Quit
    Set accessApp = Nothing
    
    WScript.Echo "  Index refreshed in " & FormatNumber(Timer - indexStart, 2) & " seconds"
    WScript.Echo "  [OK] Index complete!"
    WScript.Echo ""
    
    '───────────────────────────────────────────
    ' STEP 3: Record completion date
    '───────────────────────────────────────────
    Dim trackingFile
    trackingFile = dataDir & "LastMaintenanceDate.txt"
    
    Dim ts
    Set ts = fso.CreateTextFile(trackingFile, True) ' True = Overwrite
    ts.WriteLine FormatDateTime(Now, vbShortDate)
    ts.Close
    
    WScript.Echo "==========================================="
    WScript.Echo "[OK] Daily maintenance complete!"
    WScript.Echo "  Total time: " & FormatNumber(Timer - maintenanceStart, 1) & " seconds"
    WScript.Echo "==========================================="
    WScript.Echo ""
    
    Set fso = Nothing
End Sub


'************************
'* This function calculates hours, minutes
'* and seconds based on how many seconds
'* are passed in and returns a nice format
'************************
Public Function PrintHrMinSec(elap)
    Dim hr
    Dim min
    Dim sec
    Dim remainder
    
    elap = Int(elap) 'Just use the INTeger portion of the variable
    
    'Using "\" returns just the integer portion of a quotient
    hr = elap \ 3600 '1 hour = 3600 seconds
    remainder = elap - hr * 3600
    min = remainder \ 60
    remainder = remainder - min * 60
    sec = remainder
    
    'Prepend leading zeroes if necessary
    If Len(sec) = 1 Then sec = "0" & sec
    If Len(min) = 1 Then min = "0" & min
    
    'Only show the Hours field if it's non-zero
    If hr = 0 Then
        PrintHrMinSec = min & ":" & sec
    Else
        PrintHrMinSec = hr & ":" & min & ":" & sec
    End If
    
End Function

'====================================================================================================
'============BELOW CODE SITS WITHIN IN ACCESS DATABASE "BONYNostro.accdb"============================
'====================================================================================================

' ================================================
' STANDARD MODULE - 1EntryPoint
' ================================================

Option Compare Database

Public Const DATA_DIR_BONY       As String = "\\pc.internal.macquarie.com\FSVC\AMERICAS\wnycfsp28535" _
                                            & "\Shared\tnc\SCF\Commodity Margin Lending\CML Start-up\" _
                                            & "Middle Office Procedures\TLM_NostroRec\BONY\"

Public Const DATA_DIR_WOODSON As String = "\\pc.internal.macquarie.com\FSVC\AMERICAS\wnycfsp28535\" _
                                         & "Shared\tnc\SCF\Commodity Margin Lending\CML Start-up\" _
                                         & "Middle Office Procedures\Woodson\VBA Examples\data\"


'Public Const RAW_DATA_DIR        As String = "\\pc.internal.macquarie.com\FSVC\AMERICAS\wnycfsp28535" _
'                                            & "\Shared\tnc\SCF\Commodity Margin Lending\CML Start-up\" _
'                                            & "Middle Office Procedures\TLM_NostroRec\BONY"

Private LatestEmail As Outlook.MailItem
Private NewOutlookDataFound As Boolean

Public Sub IngestNewData(ByVal isManualUpload As Boolean, Optional ByVal Log As Scripting.TextStream)

    Dim Start As Date
    Start = Now()
    'Stop
    If Not isManualUpload Then 'New File is coming from Email
        Dim LastEmail As Outlook.MailItem
        Dim IsNewDataFound As Boolean
        Set LastEmail = InspectOutlook(IsNewDataFound)
        
        'Stop
        If Not IsNewDataFound Then
            UpdateLog Source:="Email", _
                     ValueDate:=Fix(CDBl(LastEmail.ReceivedTime)), _
                     TimeOnTask:=(Timer - Start) /86400, _
                     NewDataFound:=False, _
                     LastEmail:=LastEmail
            Exit Sub
        Else
            Debug.Print "Confirming New BONY Data found..."
            ExtractBONYDataFromHTML LastEmail
            MoveDataToImportFolder LastEmail
        End If
    End If

	Dim fso As New Scripting.FileSystemObject
	Dim Fle As Scripting.File
	For Each Fle In fso.GetFolder(DATA_DIR_BONY & "\LastImports").Files
		'Debug.Print Fle.Name
		If Fle.Type = "Text Document" Then
			Dim MetaData As TFileMetaData
			MetaData = ParseMetaData(Fle)
			'Fle.Name = Format(MetaData.ValueDate, "YYYYMMDD.txt")
			
			'Stop
			Dim LastUploadLog As DAO.Recordset
			Set LastUploadLog = GetLastUploadLog(CDbl(MetaData.ValueDate))
			
			If LastUploadLog.EOF Then
				UploadDataFromImportFolder Fle, MetaData, Log
				IsNewDataFound = True
			Else
				If MetaData.StatementRunTime - LastUploadLog("BONYRunTime").value > 0.00000001 Then
					UploadDataFromImportFolder Fle, MetaData, Log
					IsNewDataFound = True
				Else
					IsNewDataFound = False
				End If
			End If
			
			MoveDataToStorageFolder Fle, MetaData.ValueDate
			
			UpdateLog Source:=IIf(isManualUpload, "Manual", "Email"), _
					ValueDate:=CDbl(MetaData.ValueDate), _
					NewDataFound:=IsNewDataFound, _
					TimeOnTask:=(Timer - Start) / 86400, _
					LastEmail:=LastEmail, _
					BONYRunTime:=MetaData.StatementRunTime, _
					BONYLastUpdate:=MetaData.StatementLastAcctActivity
			
		End If
	Next Fle

	Debug.Print "Run Complete!! - Time Taken: " & Format(Now() - Start, "hh:mm:ss")

End Sub
	
Public Function InspectOutlook(ByRef IsNewDataFound As Boolean) As Outlook.MailItem

    Debug.Print "Searching Outlook for latest BONY email in folder '*****BONY Activity*****'"
    
    Dim LatestEmail As Outlook.MailItem
    Set LatestEmail = GetLatestHTMLFromOutlook("*****BONY Activity*****")
    Debug.Print "Retrieved Email...Subject: " & LatestEmail.Subject & " Time Received: " & LatestEmail.ReceivedTime
    
    Dim LastUpload As DAO.Recordset
    Set LastUpload = GetLastUploadLog(Fix(CDbl(LatestEmail.ReceivedTime)), True)
    
    If LastUpload.EOF Then
        IsNewDataFound = True
    Else
        'If LastUpload("SourceFileName").value Like "*_Email" Then
            If CDbl(LatestEmail.ReceivedTime) - LastUpload("MyLastEmailReceived").value > 0.00000001 Then
                IsNewDataFound = True
            Else
                IsNewDataFound = False
            End If
        'End If
    End If
    
    Set LastUpload = Nothing
    Set InspectOutlook = LatestEmail

End Function
	
Private Function GetLastUploadLog(ByVal ValueDate As Double, Optional ByVal IsEmailSourcedOnly As Boolean = False) As DAO.Recordset

    If IsEmailSourcedOnly Then
        Dim AdditionalSql As String
        AdditionalSql = " AND SourceFileName LIKE '%_Email'"
    End If
    
    Dim sql As String
    sql = "SELECT TOP 2 SourceFileName, ValueDate, MyLastUploadTime, MyLastEmailReceived, BONYRunTime, BONYLastUpdate," _
        & " UserName, TimeOnTask, NewDataFound" _
        & " FROM LastUpload" _
        & " WHERE SourceFileName Like 'BONYStatement*'" _
        & " AND ValueDate = " & ValueDate _
        & " AND NewDataFound" _
        & AdditionalSql _
        & " ORDER BY BONYLastUpdate DESC" 'returning an empty rs
    
    Set GetLastUploadLog = CurrentDb.OpenRecordset(sql)

End Function

Private Sub UpdateLog(ByVal Source As String, ByVal ValueDate As Double, ByVal NewDataFound As Boolean, ByVal TimeOnTask As Double, _
	Optional ByVal LastEmail As Outlook.MailItem = Nothing, Optional ByVal BONYRunTime As Double, Optional ByVal BONYLastUpdate As Double)
    
    Debug.Print "Logging Run..." & Format(ValueDate, "DD-MMM-YYYY")
    With GetLastUploadLog(ValueDate)
        If Not NewDataFound Then
            Dim PreviousRun As Double
            PreviousRun = !MyLastUploadTime
        End If
        
        .AddNew
        !SourceFileName = "BONYStatement_" & Source
        !ValueDate = ValueDate
        
        If LastEmail Is Nothing Then
            !MyLastEmailReceived = 0
        Else
            !MyLastEmailReceived = LastEmail.ReceivedTime
        End If
        
        !UserName = Excel.Application.UserName
        !TimeOnTask = (Timer - Start) / 86400
        !NewDataFound = NewDataFound
        
        If NewDataFound Then
            !MyLastUploadTime = Now()
            !BONYRunTime = BONYRunTime
            !BONYLastUpdate = BONYLastUpdate
        Else
            !MyLastUploadTime = PreviousRun
            !BONYRunTime = 0
            !BONYLastUpdate = 0
        End If
        
        .Update
    End With
End Sub

' ================================================
' STANDARD MODULE - 2OutlookRetrieval
' ================================================

Option Compare Database

Public Function GetLatestHTMLFromOutlook(FolderToBeSearched) As Outlook.MailItem
    
    Dim Result As Outlook.MailItem
    
    Dim objNS As Outlook.Namespace
    Dim Inbox As Outlook.MAPIFolder
    
    Set objNS = Outlook.Application.GetNamespace("MAPI")
    Set Inbox = objNS.Folders.Item("Kevin.Freehill@macquarie.com").Folders("Inbox").Folders(FolderToBeSearched)
    
    
    Dim myItems As Outlook.Items
    Set myItems = Inbox.Items
    myItems.Sort "[ReceivedTime]", True
    ' msgbox myItems(1).ReceivedTime & " > " & myItems(2).ReceivedTime, 64
    ' msgbox myItems(1).ReceivedTime > myItems(2).ReceivedTime, 64
    Dim Email As Outlook.MailItem
    For Each Email In myItems
        Debug.Print Email.Subject
        If TypeName(Email) = "MailItem" And Email.Subject = "Balance Reporting Event Standard Html Txt File" Then
            Set Result = Email
            'Debug.Print Email.Subject & vbTab & Email.ReceivedTime
            Exit For
        End If
    Next Email
    
    Set GetLatestHTMLFromOutlook = Result

End Function


Private Function GetEmailAddress() As String
    Static mapper As Scripting.Dictionary
       
    If mapper Is Nothing Then
        Set mapper = New Scripting.Dictionary
        mapper.Add "kfreehill", "Kevin.Freehill@macquarie.com"
        mapper.Add "Phil Cool", "Phil"
        mapper.Add "Artemio Colon", "Manny"
        mapper.Add "Jason Garcia", "Jason"
        mapper.Add "Alec Lesniewski", "Alec"
        mapper.Add "Shirley Ly", "Shirley"
    End If
    
    
    Set objSysInfo = CreateObject("WinNTSystemInfo")
    If mapper.Exists(objSysInfo.UserName) Then
        GetEmailAddress = mapper(objSysInfo.UserName)
    Else
        Stop
    End If

End Function

' ================================================
' STANDARD MODULE - 3DataRetrievalFromHTML
' ================================================

Option Compare Database

Private Const PWRD_TAG           As String = "<input type=""password"" name=""ACF_PASSWORD"" size=""20""><br><br>"
Private Const NEW_PWRD_TAG       As String = "<input type=""password"" name=""ACF_PASSWORD"" size=""20"" value=""scfbalances""><br><br>"

Private Const SUBMIT_TAG         As String = "<input type=""submit"" value=""Submit"">"
Private Const NEW_SUBMIT_TAG     As String = "<input id=""submit"" type=""submit"" value=""Submit"">"

Private Const BODY_TAG           As String = "</head>"

Private Declare Function ShellExecute _
    Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal Operation As String, _
    ByVal Filename As String, _
    Optional ByVal Parameters As String, _
    Optional ByVal Directory As String, _
    Optional ByVal WindowStyle As Long = vbMinimizedFocus _
    ) As Long

Public Sub ExtractBONYDataFromHTML(ByVal Email As Outlook.MailItem)

    CleanupDownloadFolder
    
    Debug.Print "Extracting BONY Data From HTML Attachment..."
    Dim EmailAttachment As Outlook.Attachment
    Set EmailAttachment = Email.Attachments(1)
    
    'Save HTML File
    With EmailAttachment
        .SaveAsFile DATA_DIR_WOODSON & "BONYFiles\" & .Filename 'Saving file with REVISED name for DB upload
    End With
    
    'Update HTML File
    Dim fso As New Scripting.FileSystemObject
    With fso.OpenTextFile(DATA_DIR_WOODSON & "BONYFiles\" & EmailAttachment.Filename, ForReading, False)
        Dim htmlText As String
        htmlText = .ReadAll
        
        Dim Result As String
        Result = Replace(htmlText, PWRD_TAG, NEW_PWRD_TAG)
        Result = Replace(Result, SUBMIT_TAG, NEW_SUBMIT_TAG)
        Result = Replace(Result, BODY_TAG, GetNewBodyTag)
    End With
    
    With fso.OpenTextFile(DATA_DIR_WOODSON & "BONYFiles\" & EmailAttachment.Filename, ForWriting, True)
        .Write Result
    End With
    
    'Extract BONY Activity Data From HTML into Downloads Directory
    OpenUrl

End Sub

Private Function GetNewBodyTag()

    GetNewBodyTag = "</head>" & vbCrLf _
        & "<script>" & vbCrLf _
        & "window.onload = function(){" & vbCrLf _
        & String(2, " ") & "document.getElementById('submit').click();" & vbCrLf _
        & "}" & vbCrLf _
        & "</script>" _
        & vbCrLf

End Function


Private Sub OpenUrl()

    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", DATA_DIR_WOODSON & "BONYFiles\EventManager.html")

End Sub

Public Sub CleanupDownloadFolder()

    Debug.Print "Cleaning up Download to Storage Folder..."
    
    Dim myDownloadsFolder As String
    myDownloadsFolder = "C:\Users\kfreehil\OneDrive - Macquarie Group\Personal Folders\Downloads"
    'Stop
    Dim fso As New Scripting.FileSystemObject
    With fso.GetFolder(myDownloadsFolder)
        Dim Fle As Scripting.File
        For Each Fle In .Files
            'Move all old downloads to PreviousDownload folder
            On Error GoTo Err_Handler
            Fle.Move myDownloadsFolder & "\PreviousDownloads\" & Fle.Name
        Next Fle
    End With
    
    Exit Sub

Err_Handler:
    If Err.Number = 58 Then
        Dim i As Integer
        i = i + 1
        Fle.Name = Split(Fle.Name, ".") (0) & "_" & i & ".csv"
        Err.Clear
        Resume
    ElseIf Err.Number = 70 Then
        Err.Clear
        Resume Next
    End If

End Sub

' ================================================
' STANDARD MODULE - 4MoveFiles
' ================================================

Option Compare Database

Private Type TLastModifiedFile
    Name As String
    LastModified As Double
End Type

Public Sub MoveDataToImportFolder(ByVal Email As Outlook.MailItem)

    'Move BONY Activity File to TLM_NostroRec\BONY\LastImports Folder
    Dim myDownloadsFolder As String
    myDownloadsFolder = "C:\Users\kfreehil\OneDrive - Macquarie Group\Personal Folders\Downloads"
    
    Dim OldFilePath As String
    'OldFilePath = FindDownload(myDownloadsFolder)
    OldFilePath = myDownloadsFolder & "\messages.txt"
    
    If Not IsDownloadComplete(OldFilePath) Then Stop
    
    Dim NewFilePath As String
    NewFilePath = DATA_DIR_BONY & "LastImports\" & Format(Email.ReceivedTime, "YYYMMDD") & ".txt"
    
    Debug.Print "Moving Download to Import Folder..."
    Dim fso As New Scripting.FileSystemObject
    With fso
        If .FileExists(NewFilePath) Then
            .DeleteFile NewFilePath
        End If
        
        .MoveFile
        Source:=OldFilePath,
        Destination:=NewFilePath
    End With

End Sub

Public Sub MoveDataToStorageFolder(ByVal Fle As Scripting.File, ByVal ValueDate As Date)

    Debug.Print "Moving Download to Storage Folder..."
    
    'Move BONY Activity File to TLM_NostroRec\BONY\ Folder
    Dim OldFilePath As String
    OldFilePath = DATA_DIR_BONY & "LastImports\" & Fle.Name
    
    Dim NewFilePath As String
    NewFilePath = DATA_DIR_BONY & Format(ValueDate, "YYYMMDD") & ".txt"
    
    Dim fso As New Scripting.FileSystemObject
    With fso
        If .FileExists(NewFilePath) Then
            .DeleteFile NewFilePath
        End If
        
        .MoveFile
        Source:=OldFilePath,
        Destination:=NewFilePath
    End With

End Sub


'Private Function FindDownload(ByVal myDownloadsFolder As String) As String
'
'
'    Dim LastDownload As TLastModifiedFile
'    LastDownload.Name = ""
'    LastDownload.LastModified = 0
'    
'    Dim fso As New Scripting.FileSystemObject
'    With fso.GetFolder(myDownloadsFolder)
'        Dim Fle As Scripting.File
'        For Each Fle In .Files
'            If Fle.Name Like "messages*" Then Stop
'            
'            If Fle.Name Like "messages*" Then
'                'Debug.Print Fle.Name & vbTab & Fle.DateLastModified
'                If Fle.DateLastModified > LastDownload.LastModified Then
'                    'Fle.Name = "BONY_Latest_Activity.txt
'                    LastDownload.Name = Fle.Path
'                    LastDownload.LastModified = Fle.DateLastModified
'                End If
'            End If
'        Next Fle
'    End With
'    
'    Debug.Print "Downloand Found: " & LastDownload.Name & vbTab & Format(LastDownload.LastModified, "YYYY-MMM-DD hh:mm:ss AM/PM")
'    FindDownload = ConfirmDownloadComplete(LastDownload.Name)
    
'End Function

'Private Function ConfirmDownloadComplete(ByVal FullPath As String) As String
'
'    Dim Ret As String
'    
'    Ret = FullPath
'    If Mid(FullPath, InStrRev(FullPath, "."), 11) = ".crdownload" Then 'If FullPath ends in .crdownload
'        Dim fso As New Scripting.FileSystemObject
'        Do
'            Debug.Print "Waiting 3 seconds for download to complete"
'            Excel.Application.Wait (Now() + TimeValue("0:00:03"))
'            
'            Loop Until Not fso.FileExists(FullPath)
'            Ret = Mid(FullPath, 1, Len(FullPath) - 11)
'    End If
'    
'    Debug.Print "Download Complete!"
'    ConfirmDownloadComplete = Ret
'    
'End Function


Private Function IsDownloadComplete(ByVal FullPath As String) As Boolean

    Dim Ret As Boolean
    Dim fso As New Scripting.FileSystemObject
    Dim i As Integer
    i = 1
    Do
        Debug.Print "Waiting 3 seconds for download to complete"
        Excel.Application.Wait (Now() + TimeValue("0:00:03"))
        i = i + 1
    Loop Until fso.FileExists(FullPath) Or i = 20
    
    If fso.FileExists(FullPath) Then
        Ret = True
        Debug.Print "Download Complete!"
    End If
    
    IsDownloadComplete = Ret

End Function


' ================================================
' STANDARD MODULE - 5ParseMeta
' ================================================

Option Compare Database

Public Type TFileMetaData
    ValueDate As Date
    StatementRunTime As Double
    StatementLastAcctActivity As Double
End Type

Public Function ParseMetaData(ByVal Fle As Scripting.File) As TFileMetaData

    Dim Ret As TFileMetaData
    Dim isComplete As Boolean
    
    With Fle.OpenAsTextStream(ForReading)
    
        Do
            Dim LineItem As String
            LineItem = Trim(.ReadLine)
            Debug.Print LineItem
            If LineItem Like "AS OF *" Then
            
                Ret.ValueDate = GetStatementValueDate(LineItem)
                
            ElseIf LineItem Like "EVENT TIME*" Or LineItem Like "TIME*" Then
            
                Ret.StatementRunTime = GetStatementRunTime(LineItem)
                
            ElseIf LineItem Like "*LAST UPDATED ON*" Then
            
                Ret.StatementLastAcctActivity = GetLastAcctActivity(LineItem)
                isComplete = True
                
            End If
        Loop Until isComplete Or .AtEndOfStream
        
    End With
    
    ParseMetaData = Ret
End Function

Private Function GetStatementValueDate(ByVal LineItem As String) As Date
    Dim Ret As Date
    
    LineItem = Right(LineItem, 10)
    Ret = DateSerial(Split(LineItem, "/")(2), Split(LineItem, "/")(0), Split(LineItem, "/")(1))
    
    GetStatementValueDate = Ret
End Function

Private Function GetStatementRunTime(ByVal LineItem As String) As Double
    Dim Ret As Variant
    Dim RunTime As String
    Dim TodaysDate As String
    
    RunTime = Split(LineItem, "TIME")(1)
    RunTime = Trim(Split(RunTime, "ET")(0))
    
    TodaysDate = Trim(Split(LineItem, "DATE:")(1))
    TodaysDate = Format(DateSerial(Split(TodaysDate, "/")(2), Split(TodaysDate, "/")(0), Split(TodaysDate, "/")(1)), "DD-MMM-YYYY")
    Ret = Format(TodaysDate & " " & RunTime, "DD-MMM-YYYY hh:mm:ss AM/PM")
    
    GetStatementRunTime = CDbl(CDate(Ret))
End Function

Private Function GetLastAcctActivity(ByVal LineItem As String) As Double
    Dim Ret As String
    Dim LatestTime As String
    Dim LatestDate As String
    
    LatestTime = Split(LineItem, " AT ")(1)
    LatestTime = Trim(Split(LatestTime, "ET")(0))
    
    LatestDate = Split(LineItem, "LAST UPDATED ON")(1)
    LatestDate = Trim(Split(LatestDate, "AT")(0))
    LatestDate = Format(DateSerial(Split(LatestDate, "/")(2), Split(LatestDate, "/")(0), Split(LatestDate, "/")(1)), "DD-MMM-YYYY")
    
    Ret = Format(LatestDate & " " & LatestTime, "DD-MMM-YYYY hh:mm:ss AM/PM")
    
    GetLastAcctActivity = CDbl(CDate(Ret))
End Function

' ================================================
' STANDARD MODULE - 6UploadToDatabase
' ================================================

Option Compare Database

Private rs As DAO.Recordset

Private iCashID As Double
Private pValueDate As Date

'**********************
'*** INGESTING DATA***
'**********************
Public Function UploadDataFromImportFolder(ByVal Fle As Scripting.File, ByRef MetaData As TFileMetaData, _
    Optional ByVal Log As Scripting.TextStream) As Boolean

    DebugPrint "Parsing New BONY Statement for Value Date..." & vbLf, Log
    'pValueDate = ParseValueDate(Fle)
    pValueDate = MetaData.ValueDate

    DebugPrint vbLf & "Clearing DB...For Value " & pValueDate, Log
    OpenAccessDBConnection pValueDate

    DebugPrint "Parsing New BONY Statement...For Value " & pValueDate, Log
    With Fle.OpenAsTextStream(ForReading)
        Dim LineItem As String
        LineItem = Trim(.ReadLine)
        
        Dim Parser As CashMovementParser
        Set Parser = New CashMovementParser
        iCashID = 1
		While Not .AtEndOfStream
			If IsNewCashMovement(LineItem) Then
				Parser.StartNew
				' Add the first line (the transaction type line)
				Parser.AddLineItem LineItem
				LineItem = Trim(.ReadLine)
				
				' Keep adding lines until we hit the start of the next movement or EOF
				While Not .AtEndOfStream And Not IsCashMovementEnd(LineItem)
					Parser.AddLineItem LineItem
					LineItem = Trim(.ReadLine)
				Wend
				
				' LineItem now contains the start of the NEXT movement (don't add it)
				Parser.ParseDetails
				AddCashItemToDB Parser
				' Do NOT read another line - LineItem already has the next transaction start
			Else
				AddNonCashItemToDB LineItem
				LineItem = Trim(.ReadLine)
			End If
			iCashID = iCashID + 1
		Wend
    End With
    
    DebugPrint "Saving New BONY Statement to DB...For Value " & pValueDate, Log

End Function

Private Sub DebugPrint(ByVal StatusUpdate As String, Optional ByVal Log As Scripting.TextStream)

    If Log Is Nothing Then
        Debug.Print StatusUpdate
    Else
        Log.WriteLine StatusUpdate
        Debug.Print StatusUpdate
    End If

End Sub

Private Function ParseValueDate(ByVal Fle As Scripting.File) As Date
    'StatementValueDate
    'StatementRunDate
    'LastAccountActivity
    
    Dim Ret As Date
    Dim blnDateFound As Boolean
    With Fle.OpenAsTextStream(ForReading)
        Do
            Dim LineItem As String
            LineItem = Trim(.ReadLine)
            Debug.Print LineItem
            If LineItem Like "AS OF *" Then
                LineItem = Right(LineItem, 10)
                Ret = DateSerial(Split(LineItem, "/")(2), Split(LineItem, "/")(0), Split(LineItem, "/")(1))
                blnDateFound = True
            ElseIf LineItem Like "EVENT TIME*" Then
            
            ElseIf LineItem Like "*LAST UPDATED ON*" Then
            
            End If
        Loop Until blnDateFound
    End With
    
    ParseValueDate = Ret
End Function


Private Function OpenAccessDBConnection(ByVal ValueDate As Date) As Boolean

    Dim sql As String
    sql = "DELETE * FROM BONYStatement"
        & " WHERE ValueDate = " & CDbl(ValueDate)
    CurrentDb.Execute sql
    
    sql = "SELECT * FROM BONYStatement"
        & " WHERE ValueDate > " & CDbl(ValueDate - 1)
    
    Set rs = CurrentDb.OpenRecordset(sql)

End Function

Private Function AddCashItemToDB(ByVal Parser As CashMovementParser) As Boolean

    With rs
        .AddNew
        !CashMovementID = iCashID
        !ValueDate = pValueDate
        !FedwireRef = Parser.ParseFedWireRef
        !CRNRef = Parser.ParseCRNRef
        !amount = Parser.ParseAmount
        !Details1 = Parser.ParseDetail1
        !Details2 = Parser.ParseDetail2
        !Details3 = Parser.ParseDetail3
        !Details4 = Parser.ParseDetail4
        !Details5 = Parser.ParseDetail5
        !Details6 = Parser.ParseDetail6
        !Details7 = Parser.ParseDetail7
        !Details8 = Parser.ParseDetail8
        !Details9 = Parser.ParseDetail9
        !Details10 = Parser.ParseDetail10
        .Update
    End With

End Function

Private Function AddNonCashItemToDB(ByVal LineItem As String) As Boolean
	
	With rs
		.AddNew
		!CashMovementID = iCashID
		!ValueDate = pValueDate
		!Details1 = Trim(LineItem)
		.Update
	End With
	
End Function

Private Function IsNewCashMovement(ByVal LineItem As String) As Boolean

    Dim Ret As Boolean
    
    If InStr(1, LineItem, "*BOOK TRANSFER CRDT") Then
        Ret = True
    ElseIf InStr(1, LineItem, "*INCOMING MONEY TRF") Then
        Ret = True
    ElseIf InStr(1, LineItem, "*ACH CREDIT RECEIVED") Then
        Ret = True
    ElseIf InStr(1, LineItem, "*LBOX DEP") Then
        Ret = True
    ElseIf InStr(1, LineItem, "*MISC SECURITY DEBIT") Then
        Ret = True
    ElseIf InStr(1, LineItem, "*OUTGOING MONEY TRAN") Then
        Ret = True
    ElseIf InStr(1, LineItem, "*ACH DEBIT RECEIVED") Then
        Ret = True
    ElseIf InStr(1, LineItem, "*MISC SECURITY CREDI") Then
        Ret = True
    ElseIf InStr(1, LineItem, "*BOOK TRANSFER DEBIT") Then
        Ret = True
    End If
    
    IsNewCashMovement = Ret

End Function

Private Function IsCashMovementEnd(ByVal LineItem As String) As Boolean

	Dim Ret As Boolean
	
	If InStr(1, LineItem, "TIME:") then
		Ret = True
	ElseIf Right(Trim(LineItem), 5) = "SALE/" then
		Ret = True
	ElseIf Trim(LineItem) Like "*END OF REPORT*" Then
		Ret = True
	End if

	IsCashMovementEnd = Ret
	
End Function


' ================================================
' CLASS MODULE - CashMovementParser
' ================================================

Private Type TCashDetails
    CashDetails As ADODB.Recordset
    ParsedDetails As Scripting.Dictionary
End Type

Private This As TCashDetails
Private CashId As Integer


Private Sub Class_Initialize()
    Set This.ParsedDetails = New Scripting.Dictionary
    Set This.CashDetails = New ADODB.Recordset
    With This.CashDetails.Fields
        .Append "CashMovementID", ADODB.DataTypeEnum.adVarWChar, DefinedSize:=20
        .Append "Details", ADODB.DataTypeEnum.adVarWChar, DefinedSize:=250
    End With
    This.CashDetails.Open
End Sub


Public Sub StartNew()
    This.ParsedDetails.RemoveAll
    With This.CashDetails
        .Filter = ""
        If Not (.BOF And .EOF) Then
            .MoveFirst
            While Not .EOF
                .Delete
                .MoveNext
            Wend
        End If
    End With
    CashId = 0
End Sub

Public Sub AddLineItem(ByVal LineItem As String)
    CashId = CashId + 1
    With This.CashDetails
        .AddNew
        !CashMovementID = CashId
        !Details = Trim(LineItem)
    End With
End Sub

Public Function IsEmpty() As Boolean
    IsEmpty = This.CashDetails.BOF And This.CashDetails.EOF
End Function

Public Function ParseFedWireRef() As String
    Dim Ret As String
    With This.CashDetails
        .Filter = "CashMovementID = 2"
        If Not (.BOF And .EOF) Then
            Ret = Trim(Split(.Fields("Details").value, " ")(0))
        End If
    End With
    ParseFedWireRef = Ret
End Function

Public Function ParseCRNRef() As String
    Dim Ret As String
    With This.CashDetails
        .Filter = "Details Like 'CRN:%'"
        If Not (.BOF And .EOF) Then
            Ret = Trim(Split(.Fields("Details").value, " ")(1))
        End If
    End With
    ParseCRNRef = Ret
End Function

Public Function ParseAmount() As Double
    Dim Ret As Variant
    With This.CashDetails
        .Filter = "CashMovementID = 1"
        Ret = Split(.Fields("Details").value, " ")
        
        Dim j As Integer
        For j = 0 To UBound(Ret) - 1
            If IsNumeric(Ret(j)) Then
                ParseAmount = CDbl(Ret(j))
                Exit For
            End If
        Next j
    End With
    'ParseAmount = ret
End Function

Public Sub ParseDetails()
    With This.CashDetails
        .Filter = ""
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Dim l As Integer
            Dim k As Integer
            m = 1
            k = 1
            While Not .EOF
                Dim Details As String
                Details = Details _
                    & .Fields("Details").value & vbCrLf
                k = k + 1
                .MoveNext
                If k > 3 Or .EOF Then
                    This.ParsedDetails.Add "Details_" & m, Left(Details, Len(Details) - 1)
                    Details = ""
                    k = 1
                    m = m + 1
                End If
            Wend
        End If
    End With
End Sub

Public Function ParseDetail1() As String
    If This.ParsedDetails.Exists("Details_1") Then
        ParseDetail1 = This.ParsedDetails("Details_1")
        'Debug.Print This.ParsedDetails("Details_1")
    End If
End Function

Public Function ParseDetail2() As String
    If This.ParsedDetails.Exists("Details_2") Then
        ParseDetail2 = This.ParsedDetails("Details_2")
        'Debug.Print This.ParsedDetails("Details_2")
    End If
End Function

Public Function ParseDetail3() As String
    If This.ParsedDetails.Exists("Details_3") Then
        ParseDetail3 = This.ParsedDetails("Details_3")
        'Debug.Print This.ParsedDetails("Details_3")
    End If
End Function

Public Function ParseDetail4() As String
    If This.ParsedDetails.Exists("Details_4") Then
        ParseDetail4 = This.ParsedDetails("Details_4")
        'Debug.Print This.ParsedDetails("Details_4")
    End If
End Function

Public Function ParseDetail5() As String
    If This.ParsedDetails.Exists("Details_5") Then
        ParseDetail5 = This.ParsedDetails("Details_5")
        'Debug.Print This.ParsedDetails("Details_5")
    End If
End Function

Public Function ParseDetail6() As String
    If This.ParsedDetails.Exists("Details_6") Then
        ParseDetail6 = This.ParsedDetails("Details_6")
        'Debug.Print This.ParsedDetails("Details_6")
    End If
End Function

Public Function ParseDetail7() As String
    If This.ParsedDetails.Exists("Details_7") Then
        ParseDetail7 = This.ParsedDetails("Details_7")
        'Debug.Print This.ParsedDetails("Details_7")
    End If
End Function

Public Function ParseDetail8() As String
    If This.ParsedDetails.Exists("Details_8") Then
        ParseDetail8 = This.ParsedDetails("Details_8")
        'Debug.Print This.ParsedDetails("Details_8")
    End If
End Function

Public Function ParseDetail9() As String
    If This.ParsedDetails.Exists("Details_9") Then
        ParseDetail9 = This.ParsedDetails("Details_9")
        'Debug.Print This.ParsedDetails("Details_9")
    End If
End Function

Public Function ParseDetail10() As String
    If This.ParsedDetails.Exists("Details_10") Then
        ParseDetail10 = This.ParsedDetails("Details_10")
        'Debug.Print This.ParsedDetails("Details_10")
    End If
End Function
