# KTM Projects in Source Control

KTM projects can be checked into source control just like any other files.  However all code and configuration in the project is inside the FPR file, so this does not lend itself well to the normal benefits of source control, such as being able to diff changes between check-ins.  One of the ways to improve this is to export out components of the project so changes in these can be tracked separately.

## Exportable Project Components

Some parts of a transformation project can be manually exported out of the FPR file.  See the next section for a programmatic solution.

### Script

Checking in a separate copy of your scripts as plain text files will allow more normal source control functionality.  This makes the code easy to view and will allow the ability to diff code changes.  You could probably also merge changes in the text files as long as you're aware that you still ultimately need to copy the merged result back into the actual project.

To manually export scripts, either copy-paste your script into files or use the "Save All Scripts" feature added to the script window's Tools menu in KTM 6.1.1, KTM 6.2.1, and KTA 7.5.

#### Script References

There is an aspect of WinWrap Basic that is worth mentioning here.  External references added to a given script sheet are stored as plain text comments at the top of the script ('#Reference {COM GUID}), however these are normally hidden in the script editing window.  These are visible in the exported script, and if you paste one into a script editing window it will disappear after your focus leaves the line, and you can confirm that the reference will now show up under Edit > References.  Thus there are no extra steps needed to restore references when pasting a script like this into the editor.

### Locators

Locator definitions can be exported, then checked in separately to make it easier to revert changes on a specific locator.  Diffing changes is not very practical since locator definitions contain a combination of binary and xml data.  The contents might give some context, but don't expect much.

To manually export a selected locator click the Design tab on the ribbon, then from the Edit group, choose Export.

### Other

Other elements have export capability and could poentially be treated the same way: Table Header Packs, Recognition Profiles, Image Cleanup Profiles, Document Review Shortcuts, and Classification Instructions (Instruction Classification phrases).

## Programmatic Export

At the end of the [KTM Script Logging Framework](DevExport/KTM Script Logging Framework.vb) you will find the function Dev_ExportScriptAndLocators and supporting functions CreateClassFolders, and ClassHierarchy.  You could call this manually during development, but what I recommend is to call it when you extract a document in Project Builder, or open a document in Validation, or something else you do frequently while developing and testing the project.  That way, the exported files are essentially updated as you work, and are easy to manually trigger if needed.  Example:

    Private Sub Document_AfterExtract(ByVal pXDoc As CASCADELib.CscXDocument)
    ' Only when run in Project Builder...
    If Project.ScriptExecutionMode=CscScriptExecutionMode.CscScriptModeServerDesign Then
        ' Update external script and locator files added to source control
        Dev_ExportScriptAndLocators()
    End If
    End Sub

## Directly Diff Project Files in Git

Wolfgang Radl explained a very cool idea to allow git to show diffs within the zipped xml of the Project FPR file (also applicable to xdc and xfd files): [Order To Chaos: Version Control and Transformations](https://www.theorycrafter.org/quipu/order-to-chaos-version-control-and-transformations/)

The two appoaches can work together: His has the natural advantage of seeing changes associated with the FPR file.  The approach here has the advantage of allowing the code and diffs to be visible even where his client side script has not been configured, and it adds the granularity of tracking locator definition changes.
