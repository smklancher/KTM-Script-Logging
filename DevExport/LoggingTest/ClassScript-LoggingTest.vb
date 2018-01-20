Option Explicit

' Class script: LoggingTest
Private Sub ValidationForm_ButtonClicked(ByVal ButtonName As String, ByVal pXDoc As CASCADELib.CscXDocument)
   On Error GoTo catch
   Select Case ButtonName
      Case "LogError"
         Err.Raise(vbObjectError, "", "Testing Error Logging")

      Case "LogMsgBox"
         Dim MsgBoxResult As VbMsgBoxResult
         MsgBoxResult = MsgBoxLog("This message, style, and user choice will be logged.", _
         vbYesNoCancel Or vbCritical, "MsgBoxLog Example", pXDoc, Nothing, _
         ValidationForm.DocViewer.ActivePageIndex + 1)

      Case "OpenLogFile"
         'Notepad is available in a directory already on the system PATH variable
         Dim TextEditorCmdLine As String
         TextEditorCmdLine = "notepad"

         'check if we are in design or runtime
         If Project.ScriptExecutionMode = CscScriptModeServerDesign Or _
            Project.ScriptExecutionMode = CscScriptModeValidationDesign Or _
            Project.ScriptExecutionMode = CscScriptModeVerificationDesign Then
            Shell(TextEditorCmdLine & " """ & Logging_CaptureLocalLogs() & Format(Now(), "yyyymmdd") & DESIGN_LOG_FILENAME & """", vbNormalFocus)
         Else
            Shell(TextEditorCmdLine & " """ & BATCH_IMAGE_LOGS & BATCH_ID_HEX & BATCH_LOG_FILENAME & """", vbNormalFocus)
         End If
   End Select

   Exit Sub

   catch:
   ErrorLog(Err, "Logging Framework Test", pXDoc, Nothing, ValidationForm.DocViewer.ActivePageIndex + 1)
   Resume Next
End Sub

Private Sub ValidationForm_DocumentLoaded(ByVal pXDoc As CASCADELib.CscXDocument)
   Dev_ExportScriptAndLocators()
End Sub
