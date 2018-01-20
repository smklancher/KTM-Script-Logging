# Kofax Transformation Modules Script Logging Framework

## Background

In 2011 I wrote a framework of logging functions for KTM script.  The [documentation](ScriptLoggingDocumentation.md) goes into some of the rationale for how it was designed.  The main goal was to make it easy to quickly drop into a project, add a few logging function calls, and know that relevant metadata would be included in the resulting logs while trying to diagnose a problem.

I also put together some thoughts on approaches to adding a KTM project to source control: [KTM Source Control](KTMSourceControl.md)

## Logging Framework Code

The [logging code](DevExport/KTM%20Script%20Logging%20Framework.vb) can be pasted into any new or existing project.  The script includes events that call the [required functions](ScriptLoggingDocumentation.md#required-functions), so if pasting into an existing project that already uses these events (Application_InitializeBatch, Batch_Open, and Batch_Close) you will need to move these calls into your existing events and remove the duplicates.

Then start adding the [useful functions](ScriptLoggingDocumentation.md#useful-functions) (ScriptLog, ErrorLog, and MsgBoxLog) as needed throughout the project to start logging.

## Testing the Logging Framework

The test project that contains the logging code has a Validation Form with the following names: LogError, LogMsgBox, and OpenLogFile. After clicking the LogError button and the LogMsgBox button, clicking the OpenLogFile button will open the newly created log file in Notepad, where something like the following has been logged:

    7/23/2011 9:38:56 PM (77936.24) -- D1\P1 (DocumentCustomer.xdc\DocumentCustomer.tif) -- ValidationDesign 1 -- [Document Routing Demo|ValidationForm_ButtonClicked# 40]
    [Error] Logging Framework Test
    1: [Document Routing Demo|ValidationForm_ButtonClicked# 40] ErrorLog Err, "Logging Framework Test", pXDoc, [Object@Nothing], ValidationForm.DocViewer().ActivePageIndex&() + 1&, False, False
    7/23/2011 9:39:04 PM (77943.87) -- D1\P1 (DocumentCustomer.xdc\DocumentCustomer.tif) -- ValidationDesign 1 -- [Document Routing Demo|ValidationForm_ButtonClicked# 19]
    MsgBox: User clicked Cancel for message "This message, MsgBox style, and user choice will be logged." (vbOkOnly, vbCritical, vbDefaultButton1, vbApplicationModal)

## Compatibility

### KTM

* This was written to be compatible with KTM 5.0 and higher.
* Some of the metadata is only available in KTM 5.5 and higher.
* The test project is compatible with KTM 5.5 and higher.

### TotalAgility

This framework was written long before KTA, and ideally it would be redesigned to better target that platform.  The following guidance is offered if using the current code on KTA:

* The way it is written should prevent it from *causing* any errors.
* You don’t need to call Logging_InitializeBatch since it won’t be able to get any of the KC info anyway
* Because the registry entries and folders for KC don’t exist, the local logs will end up written in the user’s temp directory
  * You could change the code in Logging_CaptureLocalLogs to set a static path if you want
* Because the KC XValues don’t exist, batch logs will also write to the user’s temp directory, all writing to “Unknown_KTM_Script_Batch.log”
  * Because "batches" are in the database, the concept of a batch specific log on the filesystem doesn’t fit with KTA as well.
  * You could make sure that you are always logging to the local log (or change AddToLocalLog to always be true in ScriptLog).
