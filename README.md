# Kofax Transformation Modules Script Logging Framework

## Background

While doing tech support for KTM, I often had reason to suggest that someone add logging to their project script to diagnose problems.  In some cases these projects already had logging functions, however

## Logging Framework Code

[Script Logging Documentation](ScriptLoggingDocumentation.md)

The logging code can be pasted into any new or existing project and just needs to be initialized as described in the required functions section of the documentation.  Then call the useful functions (ScriptLog, ErrorLog, and MsgBoxLog) as needed throughout the project.

## Testing the Logging Framework

The test project that contains the logging code has a Validation Form with the following names: LogError, LogMsgBox, and OpenLogFile. After clicking the LogError button and the LogMsgBox button, clicking the OpenLogFile button will open the newly created log file in Notepad, where something like the following has been logged:

    7/23/2011 9:38:56 PM (77936.24) -- D1\P1 (DocumentCustomer.xdc\DocumentCustomer.tif) -- ValidationDesign 1 -- [Document Routing Demo|ValidationForm_ButtonClicked# 40]
    [Error] Logging Framework Test
    1: [Document Routing Demo|ValidationForm_ButtonClicked# 40] ErrorLog Err, "Logging Framework Test", pXDoc, [Object@Nothing], ValidationForm.DocViewer().ActivePageIndex&() + 1&, False, False
    7/23/2011 9:39:04 PM (77943.87) -- D1\P1 (DocumentCustomer.xdc\DocumentCustomer.tif) -- ValidationDesign 1 -- [Document Routing Demo|ValidationForm_ButtonClicked# 19]
    MsgBox: User clicked Cancel for message "This message, MsgBox style, and user choice will be logged." (vbOkOnly, vbCritical, vbDefaultButton1, vbApplicationModal)

## Kofax TotalAgility

This framework was written long before KTA, and ideally it would be redesigned to better target that platform.  The following guidance is offered if using the current code on KTA:

* The way it is written should prevent it from *causing* any errors.
* You don’t need to call Logging_InitializeBatch since it won’t be able to get any of the KC info anyway
* Because the registry entries and folders for KC don’t exist, the local logs will end up written in the user’s temp directory
  * You could change the code in Logging_CaptureLocalLogs to set a static path if you want
* Because the KC XValues don’t exist, batch logs will also write to the user’s temp directory, all writing to “Unknown_KTM_Script_Batch.log”
  * Because "batches" are in the database, the concept of a batch specific log on the filesystem doesn’t fit with KTA as well.
  * You could make sure that you are always logging to the local log (or change AddToLocalLog to always be true in ScriptLog).
