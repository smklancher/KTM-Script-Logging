'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime#Scripting
'#Language "WWB-COM"
Option Explicit
'KTM Script Logging Framework
'2013-09-09 - Splitting batch logs by PID is now optional, default to off (set constant PROCESS_ID_IN_BATCHLOG_FILENAME)
'2013-08-08 - Design-time only issue in PB 6.0 fixed: When ParentFolder not known, doc/page not recorded.
'2013-05-16 - Split batch log by process ID.  Log process ID on error.
'2013-01-16 - To simplify code, introduced dependency on Scripting.FileSystemObject.
'           - Introduced dependency on Wscript.Shell to get correct path regardless of version or language of the OS.
'           - Fix for negative error numbers not logged
'2012-12-16 - Delimiter was missing, causing Windows user and module to be combined

'CONFIGURABLE LOGGING CONSTANTS

'This message will be prepended to the msgbox shown to a user if there is a script error
Public Const USER_ERROR_MSG As String = "An error has occurred in the KTM project script.  " & _
   "If this message needs to be reported to a system administrator, please take a " & _
   "screenshot with this message showing and explain what actions were taken immediately " & _
   "before the error occurred."
Public Const LOG_FILENAME As String = "_KTM_Script_Log.log"
Public Const BATCH_LOG_FILENAME As String = "_KTM_Script_Batch.log"
Public Const DESIGN_LOG_FILENAME As String = "_KTM_Script_Design.log"

Public BATCH_LOG_FULLPATH As String

'for readability messages are logged on a newline after metadata, set this to true to log to a single line
Public Const LOG_SINGLE_LINE As Boolean = False


'NONCONFIGURABLE LOGGING CONSTANTS
Public Const LOCAL_LOG As Boolean = True
Public Const BATCH_LOG As Boolean = False
Public Const IGNORE_CURRENT_FUNCTION As Integer = 1
Public Const FORCE_ERROR As Boolean = True
Public Const ONLY_ON_ERROR As Boolean = False
Public Const SUPPRESS_MSGBOX As Boolean = True

'SMK 2013-05-16 Use to get process ID in InitializeBatch
'SMK 2013-09-09 Make per process log optional, default to off
Declare Function GetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Public PROCESS_ID As Long
Public Const PROCESS_ID_IN_BATCHLOG_FILENAME As Boolean = False


'GLOBAL LOGGING VARIABLES
'This will be changed to true if it looks like we are in a Thin Client module
Public THIN_CLIENT As Boolean

Public BATCH_IMAGE_LOGS As String
Public CAPTURE_LOCAL_LOGS As String

Public BATCH_CLASS As String
Public BATCH_NAME As String
Public BATCH_ID As Long
Public BATCH_ID_HEX As String

'to support KTM 5.5 features
Public BATCH_USERID As String
Public BATCH_USERNAME As String
Public BATCH_WINDOWSUSERNAME As String
Public BATCH_USERSTRING As String 'combination of the previous three




'======== START LOGGING CODE ========

'Initialize Capture/runtime info.  Call from Application_InitializeBatch
Public Sub Logging_InitializeBatch(ByVal pXRootFolder As CscXFolder)
   'SMK 2013-05-16 - Determine process ID
   On Error Resume Next
   PROCESS_ID=GetCurrentProcessId()

   On Error GoTo CouldNotCreate
   'assume the batch log folder does not exist
   Dim LogFolderExists As Boolean
   LogFolderExists = False

   'these items are only set by Capture at runtime.  if any are set,
   ' then we are at runtime and they are all set
   If pXRootFolder.XValues.ItemExists("AC_BATCH_CLASS_NAME") Then
      'Set batchname, batchid, batch class
      BATCH_CLASS = pXRootFolder.XValues.ItemByName("AC_BATCH_CLASS_NAME").Value
      BATCH_NAME = pXRootFolder.XValues.ItemByName("AC_BATCH_NAME").Value
      BATCH_ID = CLng(pXRootFolder.XValues.ItemByName("AC_EXTERNAL_BATCHID").Value)
      BATCH_ID_HEX = Hex(BATCH_ID)

      'pad hex ID
      BATCH_ID_HEX = Right("00000000", 8 - Len(BATCH_ID_HEX)) & BATCH_ID_HEX

      'These items are only present in KTM 5.5+
      If pXRootFolder.XValues.ItemExists("AC_BATCH_WINDOWSUSERNAME") Then
         BATCH_WINDOWSUSERNAME=pXRootFolder.XValues.ItemByName("AC_BATCH_WINDOWSUSERNAME").Value
         BATCH_USERID = pXRootFolder.XValues.ItemByName("AC_BATCH_USERID").Value
         BATCH_USERNAME = pXRootFolder.XValues.ItemByName("AC_BATCH_USERNAME").Value

         'if user profiles are off these will all be the same
         If BATCH_WINDOWSUSERNAME = BATCH_USERID And BATCH_USERID = BATCH_USERNAME Then
            'user profiles is off so only use the windows user
         'SMK 2012-12-16 Delimiter was missing, causing Windows user and module to be combined
            BATCH_USERSTRING = BATCH_WINDOWSUSERNAME & " -- "
         Else
            'user profiles is on, so use all
            BATCH_USERSTRING = BATCH_WINDOWSUSERNAME & ", " & BATCH_USERNAME & _
               " (" & BATCH_USERID & ") -- "
         End If
      End If

      'set the batch logging path
      BATCH_IMAGE_LOGS = pXRootFolder.XValues.ItemByName("AC_IMAGE_DIRECTORY").Value & _
         "\" & BATCH_ID_HEX & "\Log\"

      'SMK 2013-01-16 - To simplify code, introduced dependency on Scripting.FileSystemObject.

      'To use an early bound object(FileSystemObject), add a reference to "Microsoft Scripting Runtime"
      ' C:\Windows\System32\scrrun.dll (C:\Windows\SysWOW64\scrrun.dll)
      ' Otherwise late bound object will be created via CreateObject("Scripting.FileSystemObject")
      Dim fso As Object
      Set fso = CreateObject("Scripting.FileSystemObject")
      'Dim fso As FileSystemObject

      On Error GoTo CouldNotCreate

      If Not fso.FolderExists(BATCH_IMAGE_LOGS) Then
         fso.CreateFolder(BATCH_IMAGE_LOGS)
      End If
      LogFolderExists=True

      'if creating the folder causes an error then LogFolderExists is still false
      CouldNotCreate:
      Err.Clear()
      On Error GoTo catch

      'if the folder still doesn't exist after trying to create, just use image path
      If Not LogFolderExists Then
         'we prefer to log to the "Log" folder along with the interactive modules,
         '  but if there is a problem, use the image path itself
         BATCH_IMAGE_LOGS = pXRootFolder.XValues.ItemByName("AC_IMAGE_DIRECTORY").Value & _
            "\" & BATCH_ID_HEX & "\"
      End If
   End If

   catch:
   Set fso=Nothing
End Sub


'log initial information about the batch.  Call from Batch_Open
Public Sub Logging_BatchOpen(ByVal pXRootFolder As CscXFolder)
   On Error GoTo catch

   'the project file is copied on publish and retains its original modified date
   Dim ProjectLastSave As Date
   ProjectLastSave = FileDateTime(Project.FileName)

   'We can only get the batch class publish date if "Copy project during publish" is used
   '  otherwise we will just get the project path
   Dim BatchClassPublishOrProjectPath As String

   'if the "Copy project during publish" is checked, it will be located within PubTypes\Custom
   If InStr(1, Project.FileName, "PubTypes\Custom") > 0 Then
      'with "Copy project during publish" the folder containing the project is created
      '   (thus dated) while publishing
      Dim ProjectFolder As String
      ProjectFolder = Mid(Project.FileName, 1, InStrRev(Project.FileName, "\") - 1)

      BatchClassPublishOrProjectPath = "published " & CStr(FileDateTime(ProjectFolder))
   Else
      'without "copy project during publish" the folder could have any date
      '  and the project could be anywhere, so just get the project path
      BatchClassPublishOrProjectPath = Project.FileName
   End If

   'log basics like batch name, class, id, machine name, project save date
   ScriptLog("Opening Batch """ & BATCH_NAME & """ (" & BATCH_ID & "/" & _
      BATCH_ID_HEX & ") -- " & BATCH_USERSTRING & Environ("ComputerName") & vbNewLine & _
      "Batch " & BATCH_ID & "/" & BATCH_ID_HEX & ": Batch Class """ & BATCH_CLASS & _
      """ (" & BatchClassPublishOrProjectPath & ", project saved " & CStr(ProjectLastSave) _
      & ")", LOCAL_LOG)

   Exit Sub

   'if there is an error, log it and try to keep going
   catch:
   ErrorLog(Err, "", Nothing, pXRootFolder, 0, ONLY_ON_ERROR, SUPPRESS_MSGBOX)
   Resume Next
End Sub



'Find and return the Capture\Local\Logs directory
'   caching the result to global variable CAPTURE_LOCAL_LOGS
Public Function Logging_CaptureLocalLogs() As String

   'If CAPTURE_LOCAL_LOGS is already set, just return it
   If CAPTURE_LOCAL_LOGS <> "" Then
      Logging_CaptureLocalLogs = CAPTURE_LOCAL_LOGS
      Exit Function

   Else
      On Error GoTo CouldNotCreate

      'SMK 2013-01-16 - Introduced dependency on Wscript.Shell to read "Local" path from registry.
      '                 Unlike previous method of manipulating environment path variables, this provides
      '                 the correct path regardless of version or language of the OS.

      'To use an early bound object (WshShell), add a reference to "Windows Script Host Object Model"
      ' C:\Windows\System32\wshom.ocx (C:\Windows\SysWOW64\wshom.ocx)
      ' Otherwise late bound object will be created via CreateObject("Wscript.Shell")
      Dim wsh As Object
      Set wsh = CreateObject("Wscript.Shell")
      'Dim wsh As New WshShell

      'The same registry location works on 32 or 64 bit OS (Windows redirect registry access from 32-bit apps to Wow6432Node as needed)
      CAPTURE_LOCAL_LOGS=wsh.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Kofax Image Products\Ascent Capture\3.0\LocalPath") & "\Logs\"

      'assume the batch log folder does not exist
      Dim LogFolderExists As Boolean
      LogFolderExists = False

      'SMK 2013-01-16 - To simplify code, introduced dependency on Scripting.FileSystemObject.

      'To use an early bound object(FileSystemObject), add a reference to "Microsoft Scripting Runtime"
      ' C:\Windows\System32\scrrun.dll (C:\Windows\SysWOW64\scrrun.dll)
      ' Otherwise late bound object will be created via CreateObject("Scripting.FileSystemObject")
      Dim fso As Object
      Set fso = CreateObject("Scripting.FileSystemObject")
      'Dim fso As FileSystemObject

      If Not fso.FolderExists(CAPTURE_LOCAL_LOGS) Then
         fso.CreateFolder(CAPTURE_LOCAL_LOGS)
      End If
      LogFolderExists=True

      'if creating the folder causes an error then FolderExists is still false
      CouldNotCreate:
      Err.Clear()
      On Error GoTo catch

      'if the folder still doesn't exist after trying to create, just use image path
      If Not LogFolderExists Then
         'if there is a problem getting the Capture local logs folder,
         '   log to the system temp folder
         CAPTURE_LOCAL_LOGS = Environ("Temp") & "\"
      End If

      Logging_CaptureLocalLogs = CAPTURE_LOCAL_LOGS
   End If

   catch:
      Set wsh=Nothing
      Set fso=Nothing
End Function



'primary logging function
Public Sub ScriptLog(ByVal msg As String, Optional ByVal AddToLocalLog As Boolean = False, _
   Optional ByVal pXDoc As CscXDocument = Nothing, _
   Optional ByVal pXFolder As CscXFolder = Nothing, _
   Optional ByVal PageNum As Integer = 0, _
   Optional ByVal ExtraDepth As Integer = 0)

   On Error GoTo catch

   'for readability messages are logged on a newline after metadata,
   '   but this is a matter of preference
   If Not LOG_SINGLE_LINE Then
      msg = Replace(vbNewLine & msg, vbNewLine, vbNewLine & vbTab)
   End If

   'refer to the WinWrap documentation for "CallersLine Function" for
   '   an explaination of the Depth parameter
   Dim Caller As String
   Caller = CallersLine(ExtraDepth)

   'In addition to the date/time, Timer makes it easy to see the number
   '  of seconds/hundredths of seconds between events
   '  output of Timer is also padded for readability in the log
   Dim DateString As String
   DateString = Now & " (" & Format(Timer, "00000.00") & ") -- "

   'if we have an folder/xdoc/page (optional) we can log which document for context
   Dim WhichDoc As String
   WhichDoc = Logging_IdentifyFolderDocPage(pXFolder, pXDoc, PageNum)

   If WhichDoc <> "" Then
      WhichDoc = WhichDoc & " -- "
   End If

   'current module, class/function/line logged from
   Dim ModuleAndFunction As String
   ModuleAndFunction = Logging_ExecutionModeString() & " -- " & Logging_StackLine(Caller) & " "

   ' Always print a shorter message to the intermediate pane of the script window
   ' Previously this was only done in "Design" modes, but this meant it excluded testing runtime events
   Debug.Print(Format(Timer, "00000.00") & " " & WhichDoc & ModuleAndFunction & msg)

   'check if we are in design or runtime
   If Project.ScriptExecutionMode = CscScriptModeServerDesign Or _
      Project.ScriptExecutionMode = CscScriptModeValidationDesign Or _
      Project.ScriptExecutionMode = CscScriptModeVerificationDesign Then

      'if we are in project builder, there is no "image directory" so just write to local logs
      '  Because Application_InitializeBatch is only called manually in PB,
      '  use Logging_CaptureLocalLogs() directly to ensure the path is set
      Open Logging_CaptureLocalLogs() & Format(Now(), "yyyymmdd") & DESIGN_LOG_FILENAME _
         For Append As #1
         Print #1, DateString & WhichDoc & ModuleAndFunction & msg
      Close #1
   Else
      'In case logging is attempted without or before initialization, set temp path and unknown batch
      'If Logging_InitializeBatch is called later it will correctly override these
      If BATCH_IMAGE_LOGS = "" Then
         'Keep in mind that a user has a different temp path per Windows session:
         'https://stackoverflow.com/a/6521387/221018
         BATCH_IMAGE_LOGS = Environ("Temp") & "\"
         BATCH_ID_HEX = "Unknown-" & Format(Now(), "yyyymmdd")
      End If

      'Path with or without process ID
      If PROCESS_ID_IN_BATCHLOG_FILENAME And PROCESS_ID > 0 Then
         BATCH_LOG_FULLPATH = BATCH_IMAGE_LOGS & BATCH_ID_HEX & "_" & PROCESS_ID & BATCH_LOG_FILENAME
      Else
         BATCH_LOG_FULLPATH = BATCH_IMAGE_LOGS & BATCH_ID_HEX & BATCH_LOG_FILENAME
      End If

      'During runtime, always log to the batch log
      Open BATCH_LOG_FULLPATH For Append As #1
      'batches are processed from various machines, so each line in the batch log
      '   should specify the machine
         Print #1, DateString & Environ("ComputerName") & " -- "  & BATCH_USERSTRING & _
            WhichDoc & ModuleAndFunction & msg
      Close #1

      'And add to local log if specified
      If AddToLocalLog Then
         Open Logging_CaptureLocalLogs() & Format(Now(), "yyyymmdd") & LOG_FILENAME _
            For Append As #1
         'the machine may be processing different batches/modules/users concurrently,
         '  so each line in the local log should have a batch id, userstring
            Print #1, DateString & BATCH_ID_HEX & " -- "  & BATCH_USERSTRING & _
               WhichDoc & ModuleAndFunction & msg
         Close #1
      End If
   End If

   catch:
End Sub


'Primary error logging function.  First parameter should always be "Err".
Public Sub ErrorLog(ByVal E As ErrObject, Optional ByVal ExtraInfo As String = "", _
   Optional ByVal pXDoc As CscXDocument = Nothing, _
   Optional ByVal pXFolder As CscXFolder = Nothing, _
   Optional ByVal PageNum As Integer = 0, _
   Optional ByVal ForceError As Boolean = False, _
   Optional ByVal SuppressMsgBox As Boolean = False)

   'checking if there is an error here means it does not need to be
   '   checked before the function is called
   If E = 0 And ForceError = False Then
      Exit Sub
   End If

   'ErrorMessage will be displayed To user In interactive modules
   Dim ErrorMessage As String
   ErrorMessage = "[Error] PID " & PROCESS_ID & " - " 'SMK 2013-05-16 include process ID in error

   'FIX SMK 2013-01-16 include negative numbers
   If E <> 0 Then
      ErrorMessage = ErrorMessage & E.Number & " - " & E.Description
   End If

   If ExtraInfo <> "" Then
      ErrorMessage = ErrorMessage & "  " & ExtraInfo
   End If


   'when the error handler is set it clears the error, so we must finish with the e param first
   E.Clear()
   On Error GoTo catch


   'get stack trace
   Dim Stack As String

   '1 extra depth to ignore this current function in the stack
   Stack = Logging_StackTrace(IGNORE_CURRENT_FUNCTION)

   'Add stack trace to the error message
   ErrorMessage = ErrorMessage & vbNewLine & Stack

   'log the error and stacktrace
   ScriptLog(ErrorMessage, LOCAL_LOG, pXDoc, pXFolder, PageNum, IGNORE_CURRENT_FUNCTION)


   'Display to user if not in Server or thin client
   If Project.ScriptExecutionMode <> CscScriptModeServer And Not THIN_CLIENT And _
      Not SuppressMsgBox Then
      'if the message needs to be localized, other languages can be added
      '   as seen in the script help topic:
      'Script Samples | Displaying Translated Error Messages For a Script Validation Method
      Dim LocalizedMessage As String


      Select Case Application.UILanguage
         Case "en-US"  'American English
            LocalizedMessage = USER_ERROR_MSG
         Case Else
            LocalizedMessage = USER_ERROR_MSG
      End Select

      'include info about where we are
      Dim WhichDoc As String
      WhichDoc = Logging_IdentifyFolderDocPage(pXFolder, pXDoc, PageNum)

      'include batch info if it exists
      If BATCH_ID_HEX <> "" Then
         WhichDoc = BATCH_NAME & " (" & BATCH_ID_HEX & "), Batch Class: " & BATCH_CLASS _
            & vbNewLine & WhichDoc
      End If

      MsgBox(LocalizedMessage & vbNewLine & vbNewLine & WhichDoc & vbNewLine & vbNewLine & _
         ErrorMessage, vbCritical, "Script Error")
   End If

   catch:
End Sub


'returns a stacktrace from where ever it is called
Public Function Logging_StackTrace(Optional ByVal ExtraDepth As Integer = 0) As String
   On Error GoTo catch

   Dim i As Integer
   i = ExtraDepth

   Dim CurrentStackLine As String
   CurrentStackLine = CallersLine(i)

   'as long as CallersLine returns something, stacktrace continues
   While CurrentStackLine <> ""
      'get a nicer format for the stack line
      CurrentStackLine = i & ": " & Logging_StackLine(CurrentStackLine) & _
         Mid(CurrentStackLine, InStr(1, CurrentStackLine, "]") + 1)

      'Add current line to the stack trace
      Logging_StackTrace = Logging_StackTrace & CurrentStackLine & vbNewLine

      'increment and try to get the next line (CallersLine returns blank if none)
      i = i + 1
      CurrentStackLine = CallersLine(i)

      'protect against trying to log a large stack
      If i > 10 Then
         Logging_StackTrace = Logging_StackTrace & i & ": ...Stack continues beyond " & _
            i - 1 & " frames..."
         Exit While
      End If
   Wend

   'on error exit
   catch:
End Function


'Returns string from ScriptExecutionMode Enum to indicate which module is running
Public Function Logging_ExecutionModeString() As String
   On Error GoTo catch

   'There is not currently a way to tell if a script is executing in a rich or thin client
   '  This is important because MsgBox cannot be used if we are in a thin client
   '  If a thin client is enabled for the project and we are in that module,
   '  we must assume it is a thin client
   THIN_CLIENT = False

   Select Case Project.ScriptExecutionMode
      Case CscScriptModeServer
         Logging_ExecutionModeString = "Server " & Project.ScriptExecutionInstance
      Case CscScriptModeServerDesign
         Logging_ExecutionModeString = "ServerDesign " & Project.ScriptExecutionInstance
      Case CscScriptModeUnknown
         Logging_ExecutionModeString = "Unknown"
      Case CscScriptModeValidation
         Logging_ExecutionModeString = "Validation " & Project.ScriptExecutionInstance
         If Project.WebBasedValidationEnabled Then
            THIN_CLIENT = True
         End If
      Case CscScriptModeValidationDesign
         Logging_ExecutionModeString = "ValidationDesign " & Project.ScriptExecutionInstance
      Case CscScriptModeVerification
         Logging_ExecutionModeString = "Verification"
         If Project.WebBasedVerificationEnabled Then
            THIN_CLIENT = True
         End If
      Case CscScriptModeVerificationDesign
         Logging_ExecutionModeString = "VerificationDesign"
      Case CscScriptModeDocumentReview
         Logging_ExecutionModeString = "DocumentReview"
         If Project.WebBasedDocumentReviewEnabled Then
            THIN_CLIENT = True
         End If
      Case CscScriptModeCorrection
         Logging_ExecutionModeString = "Correction"
         If Project.WebBasedCorrectionEnabled Then
            THIN_CLIENT = True
         End If
      Case Else
         Logging_ExecutionModeString = "BeyondUnknown (" & Project.ScriptExecutionMode & ")"
   End Select

   If THIN_CLIENT Then
      Logging_ExecutionModeString = Logging_ExecutionModeString & " (TC)"
   End If

   Exit Function

   catch:
   Logging_ExecutionModeString = "Unknown Module (Error " & Err.Number & ")"
End Function


'return [classname|subname#linenum]
'  input is the return of WinWrap's CallersLine function: "[macroname|subname#linenum] linetext"
'  refer to the WinWrap documentation for "CallersLine Function" regarding the Depth parameter
Public Function Logging_StackLine(ByVal Caller As String) As String
   On Error GoTo catch

   'the function name (subname) and linenum are followed by a ]
   Dim EndPos As Integer
   EndPos = InStr(Caller, "]")

   'the function name will follow a |
   Dim StartPos As Integer
   StartPos = InStrRev(Caller, "|", EndPos) + 1

   'get the function name
   Dim FunctionAndLine As String
   FunctionAndLine = Mid(Caller, StartPos, EndPos - StartPos)

   'combine with class/folder
   Logging_StackLine = "[" & Logging_SheetClass(Caller) & "|" & FunctionAndLine & "]"

   Exit Function

   catch:
   FunctionAndLine = "Unknown Function (Error " & Err.Number & ")"
End Function


'return the name of the folder or class of the script at the given depth
'  input is the return of WinWrap's CallersLine function: "[macroname|subname#linenum] linetext"
'  refer to the WinWrap documentation for "CallersLine Function" regarding the Depth parameter
Public Function Logging_SheetClass(ByVal Caller As String) As String
   On Error GoTo catch

   'the sheet name (macroname) is followed by a |
   Dim EndPos As Integer
   EndPos = InStr(Caller, "|")

   'the sheet name will follow a \ from Project Builder
   'Project Script: [C:\ProjectFolder\ScriptProject|Document_BeforeProcessXDoc#827] 'Code
   'Other Classes: [C:\1|ValidationForm_ButtonClicked# 18] 'Code
   Dim StartPosPB As Integer
   StartPosPB = InStrRev(Caller, "\", EndPos) + 1

   'the sheet name will follow a * from runtime modules
   '[*ScriptProject|Document_BeforeProcessXDoc#881] 'Code
   Dim StartPosRuntime As Integer
   StartPosRuntime = InStrRev(Caller, "*", EndPos) + 1

   'Use whichever start position is found
   Dim StartPos As Integer
   If StartPosPB > StartPosRuntime Then
      StartPos = StartPosPB
   Else
      StartPos = StartPosRuntime
   End If

   'get the sheet name
   Dim Sheet As String
   Sheet = Mid(Caller, StartPos, EndPos - StartPos)

   'numeric sheet names should be classes or folders
   If IsNumeric(Sheet) Then
      Dim SheetNum As Long
      SheetNum = CLng(Sheet)

      'sheet numbers higher than zero are classes
      If SheetNum > 0 Then
         Dim TheClass As CscClass
         Set TheClass = Project.ClassByID(SheetNum)

         'make sure the class actually exists to prevent an error accessing the name
         If Not TheClass Is Nothing Then
            Logging_SheetClass = TheClass.Name
         Else
            Logging_SheetClass = "Unknown Class (" & SheetNum & ")"
         End If
      Else
         'negative sheet numbers are folders (use absolute value for folder level)
         SheetNum = Abs(SheetNum)
         Dim TheFolder As CscFolderDef
         Set TheFolder = Project.FolderByLevel(SheetNum)

         'make sure the folder actually exists to prevent an error accessing the name
         If Not TheFolder Is Nothing Then
            Logging_SheetClass = TheFolder.Name
         Else
            Logging_SheetClass = "Unknown Folder (" & SheetNum & ")"
         End If
      End If
   ElseIf Sheet = "ScriptProject" Then
      'Project level script has the special designation "ScriptProject"
      Logging_SheetClass = "Project"
   Else
      Logging_SheetClass = "Unknown Class (" & Sheet & ")"
   End If

   Exit Function

   catch:
   Logging_SheetClass = "Unknown Class (Error " & Err.Number & ")"
End Function

'Meant to be called from Batch_Close, this will log routing, rejection, and other details
Public Sub Logging_BatchClose(ByVal pXRootFolder As CASCADELib.CscXFolder, _
   ByVal CloseMode As CASCADELib.CscBatchCloseMode)
   On Error GoTo catch

   Select Case CloseMode
      'routing is evaluated after Final, Suspend, and Error modes
      Case CscBatchCloseMode.CscBatchCloseError
         ErrorLog(Err, "Closing batch in error:" & BATCH_ID & "/" & BATCH_ID_HEX & ", " & _
            BATCH_NAME, Nothing, pXRootFolder, 0, FORCE_ERROR, SUPPRESS_MSGBOX)
         Logging_Routing(pXRootFolder)

         'find any rejected docs
         Dim RejectedMsg As String
         Logging_RejectedDocs(pXRootFolder, RejectedMsg, pXRootFolder.XValues)

         'if there are rejected docs/pages
         If RejectedMsg <> "" Then
            RejectedMsg = "The following have been rejected: " & vbNewLine & RejectedMsg
            ErrorLog(Err, RejectedMsg, Nothing, pXRootFolder, 0, FORCE_ERROR, SUPPRESS_MSGBOX)

            'Potentially take extra action if there is a script error
            '   (set by Logging_RejectedDocs)
            If pXRootFolder.XValues.ItemExists("LOGGING_SCRIPT_ERROR") Then
               'script error action
            End If
         End If

      Case CscBatchCloseMode.CscBatchCloseSuspend
         ScriptLog("Suspending Batch:" & BATCH_ID & "/" & BATCH_ID_HEX & ", " & _
            BATCH_NAME, LOCAL_LOG)
         Logging_Routing(pXRootFolder)

      Case CscBatchCloseMode.CscBatchCloseFinal
         ScriptLog("Batch Close")
         Logging_Routing(pXRootFolder)

      Case CscBatchCloseMode.CscBatchCloseParent
         'Application_InitializeBatch is not called between Child and Parent Batch_Close (SPR00093890)
         '  so initialize logging paths again otherwise Parent logging will go to the Child log
         Logging_InitializeBatch(pXRootFolder)

         'Log that we have "opened" the parent batch
         Logging_BatchOpen(pXRootFolder)
         ScriptLog("Routing complete, closing parent batch.")

      Case CscBatchCloseMode.CscBatchCloseChild
         'Note that if a child batch has been routed to a new batch class,
         '   this will Batch_Close will not fire for the child

         'Log that we have "opened" the child batch
         Logging_BatchOpen(pXRootFolder)

         'See if we can find out which tag this was created with during routing
         Dim i As Integer
         Dim BatchTag As String
         For i = 0 To pXRootFolder.XValues.Count - 1
            If Mid(pXRootFolder.XValues.ItemByIndex(i).Key, 1, _
               Len("KTM_DOCUMENTROUTING_QUEUE_")) = "KTM_DOCUMENTROUTING_QUEUE_" Then
               BatchTag = Mid(pXRootFolder.XValues.ItemByIndex(i).Key, _
                  Len("KTM_DOCUMENTROUTING_QUEUE_") + 1)
               ScriptLog("This batch has been created as a result of routing with " & _
                  "the tag: " & BatchTag)
            End If
         Next
         If BatchTag = "" Then
            ScriptLog("This batch has been created as a result of routing.")
         End If

      Case Else
         ErrorLog(Err, "Unknown Batch Close Type!", Nothing, pXRootFolder, 0, _
            FORCE_ERROR, SUPPRESS_MSGBOX)
   End Select

   Exit Sub

   'if there is an error, log it and try to keep going
   catch:
   ErrorLog(Err, "", Nothing, pXRootFolder, 0, ONLY_ON_ERROR, SUPPRESS_MSGBOX)
   Resume Next
End Sub

'Recursive function to check for rejected docs/pages, called from Logging_BatchClose
Public Sub Logging_RejectedDocs(ByVal XFolder As CscXFolder, ByRef msg As String, _
   ByRef XValues As CscXValues)
   On Error GoTo catch


   Dim i As Integer

   'recurse into folders
   For i = 0 To XFolder.Folders.Count - 1
      Logging_RejectedDocs(XFolder.Folders.ItemByIndex(i), msg, XValues)
   Next

   Dim RejectionNote As String

   'check documents
   Dim XDocInfo As CscXDocInfo
   For i = 0 To XFolder.DocInfos.Count - 1
      Set XDocInfo = XFolder.DocInfos.ItemByIndex(i)

      'check if the doc is rejected
      If XDocInfo.XValues.ItemExists("AC_REJECTED_DOCUMENT") Then
         'identify doc
         msg = msg & Logging_IdentifyFolderDocPage(Nothing, XDocInfo.XDocument)

         'add rejection note if exists
         If XDocInfo.XValues.ItemExists("AC_REJECTED_DOCUMENT_NOTE") Then
            RejectionNote = XDocInfo.XValues.ItemByName("AC_REJECTED_DOCUMENT_NOTE").Value
            msg = msg & ": " & RejectionNote & vbNewLine

            'if the rejection note mentions (S/s)cript,
            '   note for later that there has been a script error
            If InStr(1, RejectionNote, "cript") > 0 Then
               XValues.Set("LOGGING_SCRIPT_ERROR", "True")
            End If
         Else
            msg = msg & vbNewLine
         End If
      End If

      'check pages
      Dim PageIndex As Long
      For PageIndex = 0 To XDocInfo.PageCount - 1
         'check if the page is rejected
         If XDocInfo.XValues.ItemExists("AC_REJECTED_PAGE" & CStr(PageIndex + 1)) Then
            'identify page
            msg = msg & Logging_IdentifyFolderDocPage(Nothing, XDocInfo.XDocument, PageIndex + 1)

            'add rejection note if exists
            If XDocInfo.XValues.ItemExists("AC_REJECTED_PAGE_NOTE" & CStr(PageIndex + 1)) Then
               RejectionNote = XDocInfo.XValues.ItemByName("AC_REJECTED_PAGE_NOTE" & _
                  CStr(PageIndex + 1)).Value
               msg = msg & ": " & RejectionNote & vbNewLine
            Else
               msg = msg & vbNewLine
            End If
         End If
      Next

   Next

   'on error log and exit
   catch:
   ErrorLog(Err, "", Nothing, Nothing, 0, ONLY_ON_ERROR, SUPPRESS_MSGBOX)
End Sub


'Called from Logging_BatchClose to log documents that will be routed
Public Sub Logging_Routing(ByVal pXRootFolder As CscXFolder)
   On Error GoTo catch

   'This will hold the routing information we find:
   '   key=LOGGING_ROUTING_batchtag, value=folders and docs
   Dim RoutingGroups As CscXValues
   Set RoutingGroups = pXRootFolder.XValues

   'Set a flag saying all documents have been routed
   '  Finding an unrouted document will set this to false
   RoutingGroups.Set("LOGGING_ALLROUTED", "True")

   'recursively check folders for routing, adding results to RoutingGroups
   Logging_RoutingFolder(pXRootFolder, RoutingGroups)

   'If all docs are routed this batch will get deleted, so log this to local log
   If pXRootFolder.XValues.ItemByName("LOGGING_ALLROUTED").Value = "True" Then
      ScriptLog("All documents in the batch appear to be routed.  " & _
         "Batch will be deleted.", LOCAL_LOG)
   End If

   'log if the original batch will be routed to a module
   If RoutingGroups.ItemExists("KTM_DOCUMENTROUTING_QUEUE_THISBATCH") Then
      ScriptLog("This original batch will be routed to " & _
         pXRootFolder.XValues.ItemByName("KTM_DOCUMENTROUTING_QUEUE_THISBATCH").Value)
   End If

   'go through the document routing groups and log details
   Dim msg As String
   Dim BatchTag As String
   Dim i As Integer
   For i = 0 To RoutingGroups.Count - 1
      'if the XValue key begins with "LOGGING_ROUTING_"
      If Mid(RoutingGroups.ItemByIndex(i).Key, 1, _
         Len("LOGGING_ROUTING_")) = "LOGGING_ROUTING_" Then

         'the part after "LOGGING_ROUTING_"
         BatchTag = Mid(RoutingGroups.ItemByIndex(i).Key, Len("LOGGING_ROUTING_") + 1)

         msg = msg & "Routing group (" & BatchTag

         'check if it is being routed to a specific queue
         If pXRootFolder.XValues.ItemExists("KTM_DOCUMENTROUTING_QUEUE_" & BatchTag) Then
            msg = msg & ", Queue=" & _
               pXRootFolder.XValues.ItemByName("KTM_DOCUMENTROUTING_QUEUE_" & BatchTag).Value
         End If

         'check if it is being routed with a specific batch name KTM 5.5+
         If pXRootFolder.XValues.ItemExists("KTM_DOCUMENTROUTING_BATCHNAME_" & BatchTag) Then
            msg = msg & ", Batch Name=" & _
               pXRootFolder.XValues.ItemByName("KTM_DOCUMENTROUTING_BATCHNAME_" & BatchTag).Value
         End If

         'check if it is being routed to a new batch class KTM 5.5+
         If pXRootFolder.XValues.ItemExists("KTM_DOCUMENTROUTING_NEWBATCHCLASS_" & BatchTag) Then
            msg = msg & ", Batch Class=" & _
               pXRootFolder.XValues.ItemByName("KTM_DOCUMENTROUTING_NEWBATCHCLASS_" & _
               BatchTag).Value & _
               " (module will be ignored)"
         End If

         msg = msg & "): " & RoutingGroups.ItemByIndex(i).Value & vbNewLine
      End If
   Next

   'if there were any routing groups, msg won't be empty
   If msg <> "" Then
      ScriptLog(msg)
   End If


   'on error log and exit
   catch:
   ErrorLog(Err, "", Nothing, Nothing, 0, ONLY_ON_ERROR, SUPPRESS_MSGBOX)
End Sub


'recursively check folders for routed documents (or first level routed folders),
'   adding results to RoutingGroups
Public Sub Logging_RoutingFolder(ByVal XFolder As CscXFolder, ByRef RoutingGroups As CscXValues)
   On Error GoTo catch

   'only 1st level folders can be routed (but any documents can be routed)
   Dim IsFirstLevelFolder As Boolean
   IsFirstLevelFolder = False

   If XFolder.IsRootFolder = False Then 'not the root
      If XFolder.ParentFolder.IsRootFolder = True Then 'parent is the root
         IsFirstLevelFolder = True
      End If
   End If

   Dim BatchTag As String

   'check if this folder is being routed
   If IsFirstLevelFolder And XFolder.XValues.ItemExists("KTM_DOCUMENTROUTING") Then
      BatchTag = XFolder.XValues.ItemByName("KTM_DOCUMENTROUTING").Value

      Dim FolderName As String
      FolderName = Logging_IdentifyFolderDocPage(XFolder)

      'check if we've already added this group
      If RoutingGroups.ItemExists("LOGGING_ROUTING_" & BatchTag) Then
         'add this folder
         RoutingGroups.Set("LOGGING_ROUTING_" & BatchTag, _
            RoutingGroups.ItemByName("LOGGING_ROUTING_" & BatchTag).Value & "," & FolderName)
      Else
         'create it and add this document
         RoutingGroups.Set("LOGGING_ROUTING_" & BatchTag, FolderName)
      End If

      'if the folder is being routed, it will route the contents,
      '  and if routing instructions were set on these contents, they will be ignored
   Else
      'if the folder is not being routed, check if its subfolders
      Dim SubFolder As CscXFolder
      Dim i As Integer
      For i = 0 To XFolder.Folders.Count - 1
         Set SubFolder = XFolder.Folders.ItemByIndex(i)
         Logging_RoutingFolder(SubFolder, RoutingGroups)
      Next

      Dim oXDocInfo As CscXDocInfo
      Dim DocName As String

      'check for routed docs in this folder
      For i = 0 To XFolder.DocInfos.Count - 1
         Set oXDocInfo = XFolder.DocInfos.ItemByIndex(i)
         DocName = Logging_IdentifyFolderDocPage(Nothing, oXDocInfo.XDocument)

         'check if this document is being routed
         If oXDocInfo.XValues.ItemExists("KTM_DOCUMENTROUTING") Then
            BatchTag = oXDocInfo.XValues.ItemByName("KTM_DOCUMENTROUTING").Value

            'check if we've already added this group
            If RoutingGroups.ItemExists("LOGGING_ROUTING_" & BatchTag) Then
               'add this document
               RoutingGroups.Set("LOGGING_ROUTING_" & BatchTag, _
                  RoutingGroups.ItemByName("LOGGING_ROUTING_" & BatchTag).Value & "," & DocName)
            Else
               'create it and add this document
               RoutingGroups.Set("LOGGING_ROUTING_" & BatchTag, DocName)
            End If
         Else
            'If a document is not being routed, we know they are not all routed
            RoutingGroups.Set("LOGGING_ALLROUTED", "False")
         End If
      Next
   End If


   'on error log and exit
   catch:
   ErrorLog(Err, "", Nothing, Nothing, 0, ONLY_ON_ERROR, SUPPRESS_MSGBOX)
End Sub


'Identifies structure and files from Folder/Doc/Page.
Public Function Logging_IdentifyFolderDocPage(Optional ByVal XFolder As CscXFolder = Nothing, _
   Optional ByVal XDoc As CscXDocument = Nothing, _
   Optional ByVal PageNum As Integer = 0) As String
   'valid parameter combinations:
   'only folder
   'only doc, implies folder
   'doc and page number (page object doesn't link to parent doc)

   On Error GoTo catch

   'if doc was provided, use that to set folder
   If Not XDoc Is Nothing Then
      Set XFolder = XDoc.ParentFolder
   Else
      'if doc was not provided, make sure we have a folder, or exit
      If XFolder Is Nothing Then
         Exit Function
      Else
         'also exit if we only have root because we are omitting root folder info
         If XFolder.IsRootFolder Then
            Exit Function
         End If
      End If
   End If

   'Structure will show info like F#\F#\D#\P#
   Dim DocStructure As String

   'Files will show info like (#.xdc\#.tif(#))
   '  (#) is the page number in that document
   '  All folders are in Folder.xfd, so no need to add that
   Dim Files As String


   '2013-08-08 - SMK - XDoc.ParentFolder is not always set in KTM 6.0 Project Builder
   '  In that case, just skip this section
   If Not XFolder Is Nothing Then
      'get the folder structure (other than root folder since it is always there)
      Do While Not XFolder.IsRootFolder

         Set XFolder = XFolder.ParentFolder
         DocStructure = "F" & XFolder.IndexInFolder + 1 & "\" & DocStructure
         Files = Mid(XFolder.FileName, InStrRev(XFolder.FileName, "\")) & Files
      Loop
   End If

   'get the document
   If Not XDoc Is Nothing Then
      DocStructure = DocStructure & "D" & XDoc.IndexInFolder + 1
      Files = Files & Mid(XDoc.FileName, InStrRev(XDoc.FileName, "\") + 1)
   End If

   'get the page
   If PageNum > 0 And Not XDoc Is Nothing Then
      DocStructure = DocStructure & "\P" & PageNum

      Dim Page As CscXDocPage
      Set Page = XDoc.Pages.ItemByIndex(PageNum - 1)
      'make sure the page exists
      If Not Page Is Nothing Then
         Files = Files & Mid(Page.SourceFileName, InStrRev(Page.SourceFileName, "\"))
      End If
   End If

   Logging_IdentifyFolderDocPage = DocStructure & " (" & Files & ")"


   Exit Function

   catch:
   Logging_IdentifyFolderDocPage = "Unknown Folder/Doc/Page"
End Function


'Wrapper and drop-in replacement for MsgBox.
Public Function MsgBoxLog(ByVal Message As String, _
   Optional ByVal MsgType As VbMsgBoxStyle, _
   Optional ByVal Title As String, _
   Optional ByVal pXDoc As CscXDocument=Nothing, _
   Optional ByVal pXFolder As CscXFolder=Nothing, _
   Optional PageNum As Integer=0) As VbMsgBoxResult

   On Error GoTo catch

   'Figure out what kind of MsgBox style this is (one from each group)
   Dim TypeString As String

   'Buttons
   If CInt(MsgType And vbOkOnly) = vbOkOnly Then
      TypeString = "vbOkOnly, "
   ElseIf CInt(MsgType And vbOkCancel) = vbOkCancel Then
      TypeString = "vbOkCancel, "
   ElseIf CInt(MsgType And vbAbortRetryIgnore) = vbAbortRetryIgnore Then
      TypeString = "vbAbortRetryIgnore, "
   ElseIf CInt(MsgType And vbYesNoCancel) = vbYesNoCancel Then
      TypeString = "vbYesNoCancel, "
   ElseIf CInt(MsgType And vbYesNo) = vbYesNo Then
      TypeString = "vbYesNo, "
   ElseIf CInt(MsgType And vbRetryCancel) = vbRetryCancel Then
      TypeString = "vbRetryCancel, "
   End If

   'Icon
   If CInt(MsgType And vbCritical) = vbCritical Then
      TypeString = TypeString & "vbCritical, "
   ElseIf CInt(MsgType And vbQuestion) = vbQuestion Then
      TypeString = TypeString & "vbQuestion, "
   ElseIf CInt(MsgType And vbExclamation) = vbExclamation Then
      TypeString = TypeString & "vbExclamation, "
   ElseIf CInt(MsgType And vbInformation) = vbInformation Then
      TypeString = TypeString & "vbInformation, "
   End If

   'Default
   If CInt(MsgType And vbDefaultButton1) = vbDefaultButton1 Then
      TypeString = TypeString & "vbDefaultButton1, "
   ElseIf CInt(MsgType And vbDefaultButton2) = vbDefaultButton2 Then
      TypeString = TypeString & "vbDefaultButton2, "
   ElseIf CInt(MsgType And vbDefaultButton3) = vbDefaultButton3 Then
      TypeString = TypeString & "vbDefaultButton3, "
   End If

   'Default
   If CInt(MsgType And vbApplicationModal) = vbApplicationModal Then
      TypeString = TypeString & "vbApplicationModal"
   ElseIf CInt(MsgType And vbSystemModal) = vbSystemModal Then
      TypeString = TypeString & "vbSystemModal"
   ElseIf CInt(MsgType And vbMsgBoxSetForeground) = vbMsgBoxSetForeground Then
      TypeString = TypeString & "vbMsgBoxSetForeground"
   End If


   'log an error as a warning if this is used during server
   If Project.ScriptExecutionMode = CscScriptExecutionMode.CscScriptModeServer Then
      ErrorLog(Err, "Skipping a MsgBox and forcing an OK result because " & _
         "it is running during Server. Message: """ & Message & """ (" & TypeString & _
         ")", pXDoc, pXFolder, PageNum, FORCE_ERROR, SUPPRESS_MSGBOX)
      MsgBoxLog = vbOK
      Exit Function
   End If

   'MsgBox also cannot be used from a Thin Client
   If THIN_CLIENT Then
      ErrorLog(Err, "Skipping a MsgBox and forcing an OK result because " & _
         "it is running during Thin Client. Message: """ & Message & """ (" & TypeString & _
         ")", pXDoc, pXFolder, PageNum, FORCE_ERROR, SUPPRESS_MSGBOX)
      MsgBoxLog = vbOK
      Exit Function
   End If

   'show the message and grab the result
   Dim Result As VbMsgBoxResult
   Result = MsgBox(Message, MsgType, Title)

   'Find out what the user clicked
   Dim ResultString As String
   Select Case Result
      Case vbOK
         ResultString = "OK"
      Case vbCancel
         ResultString = "Cancel"
      Case vbAbort
         ResultString = "Abort"
      Case vbRetry
         ResultString = "Retry"
      Case vbIgnore
         ResultString = "Ignore"
      Case vbYes
         ResultString = "Yes"
      Case vbNo
         ResultString = "No"
      Case Else
         ResultString = "Unknown"
   End Select

   'log the details
   ScriptLog("MsgBox: User clicked " & ResultString & " for message """ & Message & """ (" & _
      TypeString & ")", BATCH_LOG, pXDoc, pXFolder, PageNum, IGNORE_CURRENT_FUNCTION)

   'return the result just like a normal MsgBox
   MsgBoxLog = Result

   Exit Function

   'if there is an error, log it and try to keep going
   catch:
   ErrorLog(Err, "", pXDoc, pXFolder, 0, ONLY_ON_ERROR, SUPPRESS_MSGBOX)
   Resume Next
End Function
'========  END   LOGGING CODE ========

'========  START LOGGING IMPLEMENTATION ========
Private Sub Application_InitializeBatch(ByVal pXRootFolder As CASCADELib.CscXFolder)
   Logging_InitializeBatch(pXRootFolder)
End Sub

Private Sub Batch_Close(ByVal pXRootFolder As CASCADELib.CscXFolder, ByVal CloseMode As CASCADELib.CscBatchCloseMode)
   Logging_BatchClose(pXRootFolder,CloseMode)
End Sub

Private Sub Batch_Open(ByVal pXRootFolder As CASCADELib.CscXFolder)
   Logging_BatchOpen(pXRootFolder)
End Sub
'========  END   LOGGING IMPLEMENTATION ========


'========  START DEV EXPORT ========
Public Function ClassHierarchy(KtmClass As CscClass) As String
   ' Given TargetClass, returns Baseclass\subclass\(etc...)\TargetClass\

   Dim CurClass As CscClass, Result As String
   Set CurClass = KtmClass

   While Not CurClass.ParentClass Is Nothing
      Result=CurClass.Name & "\" & Result
      Set CurClass = CurClass.ParentClass
   Wend
   Result=CurClass.Name & "\" & Result
   Return Result
End Function

Public Sub CreateClassFolders(ByVal BaseFolder As String, Optional KtmClass As CscClass=Nothing)
   ' Creates folders in BaseFolder matching the project class structure

   Dim SubClasses As CscClasses
   If KtmClass Is Nothing Then
      ' Start with the project class, but don't create a folder
      Set KtmClass = Project.RootClass
      Set SubClasses = Project.BaseClasses
   Else
      ' Create folder for this class and become the new base folder
      Dim fso As New Scripting.FileSystemObject, NewBase As String
      BaseFolder=fso.BuildPath(BaseFolder,KtmClass.Name)
      If Not fso.FolderExists(BaseFolder) Then
         fso.CreateFolder(BaseFolder)
      End If
      Set SubClasses = KtmClass.SubClasses
   End If

   ' Subclasses
   Dim ClassIndex As Long
   For ClassIndex=1 To SubClasses.Count
      CreateClassFolders(BaseFolder, SubClasses.ItemByIndex(ClassIndex))
   Next
End Sub



Public Sub Dev_ExportScriptAndLocators()
   ' Exports design info (script, locators) to to folders matching the project class structure
   ' Default to \ProjectFolderParent\DevExport\(Class Folders)
   ' Set script variable Dev-Export-BaseFolder to path to override
   ' Set script variable Dev-Export-CopyName-(ClassName) to save a separate named copy of a class script

   ' Make sure you've added the Microsoft Scripting Runtime reference
   Dim fso As New Scripting.FileSystemObject
   Dim ExportFolder As String, ScriptFolder As String, LocatorFolder As String

   ' Either use the provided path or default to the parent of the project folder
   If fso.FolderExists(Project.ScriptVariables("Dev-Export-BaseFolder")) Then
      ExportFolder=Project.ScriptVariables("Dev-Export-BaseFolder")
   Else
      ExportFolder=fso.GetFile(Project.FileName).ParentFolder.ParentFolder.Path & "\DevExport"
   End If

   ' Create folder structure for project classes
   If Not fso.FolderExists(ExportFolder) Then fso.CreateFolder(ExportFolder)
   CreateClassFolders(ExportFolder)

   ' Here we use class index -1 to represent the special case of the project class
   Dim ClassIndex As Long
   For ClassIndex=-1 To Project.ClassCount-1
      Dim KtmClass As CscClass, ClassName As String, ScriptCode As String, ClassPath As String

      ' Get the script of this class
      If ClassIndex=-1 Then
         Set KtmClass=Project.RootClass
         ScriptCode=Project.ScriptCode
      Else
         Set KtmClass=Project.ClassByIndex(ClassIndex)
         ScriptCode=KtmClass.ScriptCode
      End If

      ' TODO: check if script is "empty": Option Explicit \n\n ' Class script: {classname}

      ' Get the name and file path for the class
      ClassPath = fso.BuildPath(ExportFolder, ClassHierarchy(KtmClass))
      ClassName=IIf(ClassIndex=-1,"Project",KtmClass.Name)

      ' TODO: Possibly change to match the naming conventions used in the KTM 6.1.1+ feature to save scripts.

      ' Export script to file
      Dim ScriptFile As TextStream
      Set ScriptFile=fso.CreateTextFile(ClassPath & "\ClassScript-" & ClassName & ".vb",True,False)
      ScriptFile.Write(ScriptCode)
      ScriptFile.Close()

      ' Save a copy if a name is defined
      Dim CopyName As String
      CopyName=Project.ScriptVariables("Dev-Export-CopyName-" & ClassName)

      If Not CopyName="" Then
         Set ScriptFile=fso.CreateTextFile(ClassPath & "\" & CopyName & ".vb",True,False)
         ScriptFile.Write(ScriptCode)
         ScriptFile.Close()
      End If

      ' Export locators (same as from Project Builder menus)
      Dim FileName As String
      Dim LocatorIndex As Integer
      For LocatorIndex=0 To KtmClass.Locators.Count-1
         If Not KtmClass.Locators.ItemByIndex(LocatorIndex).LocatorMethod Is Nothing Then
            FileName="\" & ClassName & "-" & KtmClass.Locators.ItemByIndex(LocatorIndex).Name & ".loc"
            KtmClass.Locators.ItemByIndex(LocatorIndex).ExportLocatorMethod(ClassPath & FileName, ClassPath)
         End If
      Next
   Next
End Sub
'========  END   DEV EXPORT ========


Private Sub Application_InitializeScript()
   'Test logging before Logging_InitializeBatch has been called to set the batch image path
   ScriptLog("This will be logged to temp folder")
End Sub
