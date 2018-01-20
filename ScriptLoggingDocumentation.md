# KTM Script Logging Documentation

## Log File Location

When choosing a log file location, consider that it must be accessible from all of the machines and user accounts running each of the KTM modules, both interactive and automatic.  Because of this, the best choices are likely to be locations to which KTM modules already writes its own log files.  There are two locations where KTM modules write log files:

### [Batch Image Folder]\Log

This log location is located alongside the images for the batch, where we will use the naming convention [Batch ID]\_KTM\_Script\_Batch.log.  Interactive KTM modules and some Capture modules already write their logs to this location.  When the batch is released, these logs are deleted.  This has the benefit of allowing us to log a generous amount of information here and not worry that logs will take too much space or become too large to analyze.  But for this reason, we want to make sure that any errors or other notable events are _also_ logged to a permanent location.  Additionally, there is no Batch Image Folder during design in Project Builder.

### Kofax Capture Local Logs folder

KTM Server writes daily logs to the Capture local logs folder (C:\ProgramData\Kofax\Capture\Local\Logs) and so we will do the same using the naming convention YYYMMDD\_KTM\_Script\_Log.log.  This folder exists on each machine and we will use it to log errors and other information important enough that we want it to persist even if a batch releases.  During design in Project Builder, we will log both local _and batch log_ messages to YYYMMDD\_KTM\_Script\_Design.log

Additionally while running during design, we will use Debug.Print to log any messages to the Immediate Window of the script editor.

## Logging Metadata

Every message logged with these logging functions includes useful information.  Consider the following example:

    7/15/2011 5:11:00 PM (61859.86) -- MachineName -- 0000000A -- WindowsUser, Capture User Name (CaptureUserID) – F1\D1\P1 (1.xdc,1.tif) -- Server 1 -- [Project|Logging\_BatchOpen#226]
    Message

Much like a news article, we want our log to answer "Who, What, When, Where?"  If a problem occurs, we will be able to use this information to help determine the "How?" and the "Why?"

* **Who: WindowsUser, Capture User Name (CaptureUserID)** - These properties are only exposed in KTM 5.5, so they will not be logged in KTM 5.0.  If User Profiles is not enabled, then only WindowsUser will be logged.
* **What: Message** - This is the message that we are logging.
* **When: Date Time (Seconds/Hundredths of Seconds)** - Date and time are standard in any log, but seconds/hundredths allow for quickly seeing duration between logged events.
* **Where:**
  * **MachineName** - The name of the system where the code is running.  Because all of the statements sent to the local log will be on the same machine, the machine name is only logged in the batch log.
  * **0000000A** - The hex ID of the batch being processed.  Because a batch log will always be about the same batch, the batch ID is only logged to the local log.  The local log may show multiple batches processing concurrently, so it is important that the batch ID be present in each line.
  * **F1\D1\P1 (1.xdc,1.tif)** - This is the Folder/Document/Page identification.  The example above indicates that we are dealing with the first page in the first document in the first folder, with 1.xdc storing the document and 1.tif storing the page.  The root folder is always present and always stored in Folder.xfd, so that information is omitted.
  * **Server 1** - The module and instance.
  * **[Project|Logging\_BatchOpen#226]** - This indicates the class, function, and line of script from which the message is being logged.  The class can be an important piece of context because some script events can be defined at multiple levels in a class hierarchy.

The primary benefit of using a logging framework that includes this metadata by default is that it frees us from having to manually include these common pieces of information.  While troubleshooting a For Loop without this framework we might log a message like "Function X, Document #, Page #, Beginning of Loop", whereas with the framework, we could just log "Beginning of Loop" and know that the other data and more will already be included in the log.

## Logging Functions

The functions of the logging framework are provided in full at the end of this document and are ready to be copy-pasted into a project.  The following sections briefly describe each function and in some cases explain the rationale for how they are written.  The only requirement is to follow the instructions in the Required Functions section.  The section on Useful Functions explains the functions that we intend to use to log messages throughout our project.  The Supporting Functions are used by the Useful Functions and are generally not used directly.

### Required Functions

The only absolute requirement is to call Logging\_InitializeBatch as described below, but it is also strongly recommended to use Logging\_BatchOpen and Logging\_BatchClose as described in this section.

#### Logging\_InitializeBatch

The only absolute requirement to using this framework of logging functions is that the function Logging\_InitializeBatch is called from the Application\_InitializeBatch event.  This will initialize all of the batch-specific information including the location that we will use for our batch log.

It is imperative that it is called specifically from the Application\_InitializeBatch event rather than one of the other early events.   Application\_InitializeScript does not have access to batch-specific information and Batch\_Open is only fired once per batch even if the batch was opened by multiple extraction processes which each needed to be initialized.  The required initialization should look like this:

    Private Sub Application_InitializeBatch(ByVal pXRootFolder As CASCADELib.CscXFolder)
        Logging_InitializeBatch(pXRootFolder)
    End Sub

This should come before any other code in the Application\_InitializeBatch event, because any logging which is done before this initialization, including any logging in the Application\_InitializeScript event, will not have the benefit of batch-specific information.  This means that, while the local log will work normally, the batch log location will not have been initialized and will fall back to the system's temp directory.  This will not cause any problems, but is just a limitation to keep in mind.

#### Logging\_BatchOpen

Logging\_BatchOpen should be called at the beginning of the Batch\_Open event like so:

    Private Sub Batch_Open(ByVal pXRootFolder As CASCADELib.CscXFolder)
        Logging_BatchOpen(pXRootFolder)
    End Sub

This will log a message like the following to both the batch and local logs:

    Batch 789/00000315: Batch Class "Routing" (published 7/18/2011 9:51:57 AM, project saved 7/18/2011 9:51:45 AM)

This clearly establishes not only the batch and batch class, but also the dates that the batch class and KTM Project used by the batch were last modified.  This can prevent confusion if a batch being analyzed was created with a previous version of a batch class.  Seeing different dates written to the same batch log can also establish if the "Update Batch Class" feature was used on the batch.

#### Logging\_BatchClose

Logging\_BatchClose should be called at the end of the Batch\_Close event like so:

    Private Sub Batch_Close(ByVal pXRootFolder As CASCADELib.CscXFolder, ByVal CloseMode As
    CASCADELib.CscBatchCloseMode)
        Logging_BatchClose(pXRootFolder,CloseMode)
    End Sub

If the batch is closing in Error mode, then the batch will be checked for rejected documents with the function Logging\_RejectedDocs to log details of any rejected documents.  See that function's description for more details.

If the batch is closing in Error mode, Suspend mode, or Final ("normal") mode, then the batch will be checked by Logging\_Routing to log details of documents or folders that will be routed.  See that function's description for more details.

If the batch is closing in Child mode, created from routing, this function will log the tag that the routing group used.

### Useful Functions

There are three functions that encompass the majority of the functionality of this framework and that we should use throughout our code.  ScriptLog is the main logging function and benefits from the metadata discussed in earlier in this document.  ErrorLog is used to log Err information and stack trace directly from the Err object.  Finally, MsgBoxLog is a wrapper and drop-in replacement for the MsgBox function which not only logs the message, but also logs the user's response and suppresses the dialog in non-interactive contexts.

#### ScriptLog

ScriptLog is the function we will use to log messages with the benefits of all the metadata described in the Logging Metadata section.  Here is the function definition:

    Public Sub ScriptLog(ByVal msg As String, Optional AddToLocalLog As Boolean=False, _
        Optional ByVal pXDoc As CscXDocument=Nothing, Optional ByVal pXFolder As CscXFolder=Nothing, _
        Optional PageNum As Integer=0, Optional ExtraDepth As Integer=0)

_msg_ - The only required parameter is the message that you would like to log.

_AddToLocalLog_ - The way this logging framework is constructed is that _all_ messages are logged to the batch log and then this optional parameter specifies if a particular message should also be logged to the local log.  Thus, you can think of the batch log as a verbose log and the local log as the error/important message log.  For clarity, it is recommended to specify this parameter using the defined constants LOCAL\_LOG or BATCH\_LOG instead of literal True/False values.

_pXDoc_, _pXFolder_, _PageNum_ – These optional parameters are passed to Logging\_IdentifyFolderDocPage so we can log identifying information about the Folder/Doc/Page.  See the description of that function in the Supporting Functions section for more details.  Many functions or loops within functions in KTM script operate on a particular Folder, Document, or Page.  By providing that context in our log, it will be great assistance in allowing us to trace messages or errors to specific documents or files.

*ExtraDepth* – This changes the context of the function on the stack from which we are logging.  This is only useful if we are writing other logging related functions.  For example, our ErrorLog function adds extra depth so that in the log, instead of saying it is logging a message from "ErrorLog", it will say it is logging from the function that called ErrorLog.  If this parameter is to be used, it is recommended for the sake of clarity to use the defined constant IGNORE\_CURRENT\_FUNCTION instead of a literal integer.

#### ErrorLog

ErrorLog allows us to log errors in a more organized fashion than simply logging the string from Err.Description.  It also logs a stack trace and shows a message box when used from interactive modules.  Here is the function definition:

    Public Sub ErrorLog(ByVal E As ErrObject, Optional ByVal ExtraInfo As String="", _
        Optional ByVal pXDoc As CscXDocument=Nothing, Optional ByVal pXFolder As CscXFolder=Nothing, _
        Optional PageNum As Integer=0, Optional ForceError As Boolean=False, _
        Optional SuppressMsgBox As Boolean=False)

_E_ – This parameter should always be called by passing the "Err" object which already exists as part of WinWrap's error handling.  We will let the ErrorLog function itself check if there is actually an error to log, which reduces the amount of code that goes into error handlers that are set in every function.  For example, without having to manually check if there is an error or placing an Exit Sub above the error handler, this block of code would still only log an error if one occurs:

    Public Sub TestErrorHandler()
        On Error GoTo catch
        'Code that might error
        catch:
        ErrorLog(Err) 'Only logs if there is an error
    End Sub

From the Err object we will log the error number and description.  Notably, the Err object does not tell us what line on which the error occurred, so it is important to be aware that the top line of our stack trace just references the line that calls the ErrorLog function.

_ExtraInfo_ – This parameter allows you to pass information in addition to the information already built into the Err object.  When including extra information via this parameter, be mindful of the metadata which will already be included so that you do not spend effort including duplicate information.  This function is useful in combination with the ForceError parameter.

_pXDoc_,  _pXFolder_, _PageNum_ – These optional parameters are passed to ScriptLog and on to Logging\_IdentifyFolderDocPage so we can log identifying information about the Folder/Doc/Page.  See the description of those functions for more details.

_ForceError_ – The normal operation of this function is to only log an error if one is reflected in the Err object, but we can force it to log an error with this parameter.  If an error is forced using this parameter then we should make sure to put something informative in the ExtraInfo parameter.  For clarity it is recommended to use the defined constants FORCE\_ERROR or ONLY\_ON\_ERROR instead of literal True/False values.

_SuppressMsgBox_ – When the ErrorLog function is used from interactive modules that are not thin clients, it will show a message box with the error, stack trace, and instructions about reporting the error to an Administrator if needed.  The constant USER\_ERROR\_MSG defines the instructions to the user and defaults to the message seen in the screenshot below.  Though only an English message is currently defined, the basic structure in the function is already built to accommodate localized messages.  If we do not want our error to show a message box in interactive modules, we can pass the constant SUPPRESS\_MSGBOX.

#### MsgBoxLog

MsgBoxLog is a wrapper and drop-in replacement for the MsgBox function which not only logs the message, but also logs the user's response and suppresses the dialog in non-interactive contexts. Here is the function definition:

    Public Function MsgBoxLog(ByVal Message As String, Optional ByVal MsgType As VbMsgBoxStyle, _
        Optional ByVal Title As String, Optional ByVal pXDoc As CscXDocument=Nothing, _
        Optional ByVal pXFolder As CscXFolder=Nothing,
        Optional PageNum As Integer=0) As VbMsgBoxResult

_Message_, _MsgType_, _Title, MsgBoxLog_ – Because this is a drop-in replacement for WinWrap's MsgBox function, these three parameters, and the return value of the function, are exactly the same.

_pXDoc_, _pXFolder_, _PageNum_ – These optional parameters are passed to ScriptLog and on to Logging\_IdentifyFolderDocPage so we can identify information about the Folder/Doc/Page.  See the description of those functions for more details.

Trying to display a MsgBox during KTM Server or while running a thin client may halt the module without a user able to click the message.  As such, this function will suppress the dialog and return a vbOK result instead of waiting for user input.  Notably, if the thin client is enabled for a particular module, but we open a batch in the rich client, this will cause the MsgBox to be suppressed.  See the explanation of the supporting function, Logging\_ExecutionModeString, for details.

An example of what might be logged from this function is as follows:

    MsgBox: User clicked Yes for message "This message, MsgBox style, and user choice will be logged." (vbOkOnly, vbCritical, vbDefaultButton1, vbApplicationModal)

### Supporting Functions

These supporting functions are used by the primary functions in the framework and are not generally called directly, but familiarity with how they work may still be useful.

#### Logging_CaptureLocalLogs

Logging\_CaptureLocalLogs determines the local logging path, creates it if needed, and then caches the path for future calls.  If there is any problem, we will fall back to the system's temp path.

#### Logging_ExecutionModeString

Logging\_ExecutionModeString is used to return a string describing what module the script is being run from.  There is not currently a way to tell if a script is executing in a rich or thin client.  Because of this, we check if the thin client is enabled in the project for a particular module.  Then if code is running from that module, we have to assume that it is running in the thin client.  This appends "(TC)" to the module name and sets the global Boolean variable THIN\_CLIENT to True so that other functions can adjust behavior accordingly.

#### Logging_IdentifyFolderDocPage

Logging\_IdentifyFolderDocPage will identify a Folder/Document/Page in terms of both structure and files.  For example "F1\D1\P1 (1.xdc,1.tif)" indicates that we are dealing with the first page in the first document in the first folder, with 1.xdc storing the document and 1.tif storing the page.  The root folder is always present and always stored in Folder.xfd, so that information is omitted.  The parameter combinations we can use with this function are just a folder, just a document (from which we determine the folder), or a document and a page number.  If either the document or the folder are not being included, the parameter can be passed with the keyword "Nothing".  If the page is not being included, the parameter can be passed as 0.  The page number is one-based as opposed to many PageIndex properties in KTM which are zero-based.

#### Logging_RejectedDocs

Logging\_RejectedDocs is a recursive function that goes into a folder and checks its documents and pages to see if any have been rejected.  This function is called from Logging\_BatchClose when a batch is closed in error and if any rejected documents are found it will log a message like this:

[Error]   The following have been rejected:
D1\P1 (1.xdc\1.tif): Test page rejection note.
D2 (2.xdc): Test document rejection note.

#### Logging_RoutingFolder

Logging\_RoutingFolder is a recursive function that goes into folders and checks for folders or documents that have XValues set such that they will be routed. Routing folders was introduced in KTM 5.5.  Only first level folders under the root folder can be routed, and doing so implies routing all of its contents. The information collected by this function is used by Logging\_Routing.

#### Logging_Routing

Logging\_Routing uses Logging\_RoutingFolder to gather information about documents and folders being routed. Using this information, it will log information useful in understanding what will be routed and where.  This accounts for features in KTM 5.0 such as routing specific documents or the parent batch to a separate queue.  It also accounts for features introduced in KTM 5.5 such as routing folders, routing to a new batch class, or naming the newly created batch.  Notably, if all documents in the batch are being routed, it logs a message to the local log noting that this batch will be deleted because all of its contents have been routed.  This function is called from Logging\_BatchClose and if any routed folders or documents are found it will log a message like this:

    Routing group (Test, Queue=KTM.Validation3): D1 (1.xdc)

#### Logging_SheetClass

Logging\_SheetClass works on the return value of WinWrap's CallersLine function which, according to WinWrap documentation, has the format "[macroname|subname#linenum] linetext".  The problem is that the "macroname" or "sheet" of script is just a number when we run CallersLine in KTM.  This isn't very useful to us, so this function matches the number to the Class/FolderClass and returns the name.  A notable aspect of doing this is that we can distinguish between events with the same name defined in different classes.

#### Logging_StackLine

Logging\_StackLine combines the class name from Logging\_SheetClass with the function name and line number information from WinWrap's CallersLine function.

#### Logging_StackTrace

Logging\_StackTrace returns a stack trace by calling Logging\_StackLine with successively deeper levels of WinWrap's CallersLine function.  It is called by ErrorLog to generate stack traces where errors occur.