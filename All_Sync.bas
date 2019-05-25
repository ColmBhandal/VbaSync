Attribute VB_Name = "All_Sync"
Option Explicit

Const EH_PREFIX = "EH_"
Private Const WB_PREFIX = "Workbook_"
Private Const CONFIG_FILE_NAME = "VbaSync.config"
Private Const SPECIFIC_SYNCLIST_FILE_NAME = "specificSyncList.config"
Private Const MISC_SYNCLIST_FILE_NAME = "genSyncList.config"
Private Const MISC_REL_KEY = "miscRel: "
Private Const MISC_ABS_KEY = "miscAbs: "
Private Const SPECIFIC_REL_KEY = "specificRel: "
Private Const SPECIFIC_ABS_KEY = "specificAbs: "
Private Const MISC_SYNCLIST_REL_KEY = "miscListRel: "
Private Const MISC_SYNCLIST_ABS_KEY = "miscListAbs: "
'Forms can be annoying- coming up as diffs all the time, so you can use this to turn export off for them
Private Const EXPORT_FORMS = True


Public Sub ImportModulesWarn()
    Dim answer As Integer
    answer = MsgBox("Are you sure you want to proceed?" & vbCrLf & _
    "Import will overwrite the following modules with data from disk: " & _
    vbCrLf & Join(specificWhiteList(), ",") & vbCrLf & Join(miscWhiteList(), ",") & vbCrLf, _
    vbYesNo + vbQuestion, "Import and Override?")
    If answer = vbNo Then
        Debug.Print "!!!!!!! No Import done. User cancelled."
    Else
        Call ImportModules
    End If
End Sub

Public Sub ExportModules()
    Dim lOrphans As String
    lOrphans = mGetOrphanedModulesAsString()
    If lOrphans <> "" Then
        Call MsgBox("Warning! Orphaned modules listed below won't be exported." _
            & lOrphans, vbExclamation)
    End If
    Debug.Print "***** Exports starting"
    ExportMiscModules
    ExportProjectSpecificModules
    Debug.Print "***** All exports complete"
End Sub

Private Sub ExportProjectSpecificModules()
    Dim exportFolder As String: exportFolder = createFolderWithProjectSpecificVBAFiles
    Dim whiteList() As String: whiteList = specificWhiteList()
    Call ExportModulesTargeted(exportFolder, whiteList)
End Sub

Private Sub ExportMiscModules()
    Dim exportFolder As String: exportFolder = createFolderWithVBAMiscFiles
    Dim whiteList() As String: whiteList = miscWhiteList()
    Call ExportModulesTargeted(exportFolder, whiteList)
End Sub

Private Sub ExportModulesTargeted(exportFolder As String, whiteList() As String)
Attribute ExportModulesTargeted.VB_ProcData.VB_Invoke_Func = "p\n14"
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    Debug.Print "***** Ready for export to: " & exportFolder

    Set wkbSource = ThisWorkbook
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        " not possible to export the code"
    Exit Sub
    End If
    
    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If exportFolder = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    szExportPath = exportFolder & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        If (Not isWhiteListed(szFileName, whiteList)) Then
            bExport = False
        ElseIf (importImpossible(cmpComponent)) Then
            Dim msgTrap As VbMsgBoxResult
            msgTrap = MsgBox("The module " & szFileName & " will be impossible to import." _
            & vbCrLf & "Would you like to export it anyway without a file extension?", vbYesNo)
            Select Case msgTrap
                Case vbNo
                    bExport = False
            End Select
        End If

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
                Dim shouldExportForms As Boolean
                shouldExportForms = EXPORT_FORMS
                If Not shouldExportForms Then bExport = False
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
        End Select
        
        If bExport Then
            Dim exportName As String: exportName = szExportPath & szFileName
            'Try kill the file if it exists, else just skip
            On Error Resume Next
            Kill exportName
            On Error GoTo 0
            cmpComponent.Export exportName
            Debug.Print "Exported " & szFileName
        End If
   
    Next cmpComponent
    
    Debug.Print "***** Completed export to: " & exportFolder
End Sub

Function importImpossible(cmpComponent As VBIDE.VBComponent) As Boolean
    Select Case cmpComponent.Type
        Case vbext_ct_Document
            importImpossible = True
        Case Else
            importImpossible = False
    End Select
End Function

Private Sub ImportModules()
    Call mWarnForOrphans
    Debug.Print "----- Imports starting"
    'Need to do specific first or import fails
    Call ImportProjectSpecificModules
    Call ImportMiscModules
    'BELOW IS AN EXTREMELY IMPMORTANT ASPECT OF THE IMPORT PROCESS
    'The below call must be asynchronous, so that this module can die
    'along with all those modules feeding it, and their new versions can be renamed.Dim procCallWithParams As String
    Dim procCallWithParams As String
    Application.OnTime Now + TimeSerial(0, 0, 1), "renameNumberSuffixedComponents"
    Debug.Print "----- All Imports complete"
End Sub

Private Sub mWarnForOrphans()
    Dim lOrphans As String
    lOrphans = mGetOrphanedModulesAsString()
    If lOrphans <> "" Then
        Debug.Print ("!!!!!!!!! Orphaned modules exist: " & lOrphans)
    End If
End Sub

Public Function mGetOrphanedModulesAsString() As String
    Dim lOrphan As String
    Dim loopVar As Variant
    Dim lResult As String
    lResult = ""
    For Each loopVar In mGetOrphanedModuleNames()
        lOrphan = loopVar
        lResult = lResult & vbNewLine & lOrphan
    Next
    mGetOrphanedModulesAsString = lResult
End Function

Private Function mGetOrphanedModuleNames() As Collection
    Dim resultComponents As New Collection
    Dim wkbTarget As Excel.Workbook: Set wkbTarget = ThisWorkbook
    Dim vbProj As VBIDE.VBProject
    Set vbProj = wkbTarget.VBProject
    Dim cmpComponents As VBIDE.VBComponents
    Set cmpComponents = vbProj.VBComponents
    Dim vbComp As VBIDE.VBComponent
    
    For Each vbComp In cmpComponents
        Dim compName As String: compName = vbComp.Name
        Dim listNameVar As Variant
        Dim listName As String
        If mIsOrphanedName(compName) And _
            vbComp.Type <> vbext_ct_Document Then
            resultComponents.Add vbComp.Name
        End If
    Next
    Set mGetOrphanedModuleNames = resultComponents
End Function

Private Function mIsOrphanedName(pName As String) As Boolean
    Dim lResult As Boolean
    lResult = Not isWhiteListed(pName, specificWhiteList())
    lResult = lResult And Not isWhiteListed(pName, miscWhiteList())
    'Don't consider the special sync module as an orphan- it's cleaner if it's left out of import/export.
    If pName = "All_Sync" Then lResult = False
    mIsOrphanedName = lResult
End Function

Private Sub ImportProjectSpecificModules()
    Call ImportModulesTargeted(createFolderWithProjectSpecificVBAFiles, specificWhiteList)
End Sub

Private Sub ImportMiscModules()
    Call ImportModulesTargeted(createFolderWithVBAMiscFiles, miscWhiteList)
End Sub

Private Sub ImportModulesTargeted(importFolder As String, whiteList() As String)
    Dim wkbTarget As Excel.Workbook: Set wkbTarget = ThisWorkbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.file
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    'Get the path to the folder with modules
    If importFolder = "Error" Then
        MsgBox "Problem with import folder. Quitting."
        Exit Sub
    End If

    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = importFolder & "\"
    Debug.Print "Ready to import files from: " & szImportPath
            
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       Debug.Print "There are no files to import"
       Exit Sub
    End If
    
    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    Dim vbProj As VBIDE.VBProject
    Set vbProj = wkbTarget.VBProject
    
    For Each objFile In objFSO.GetFolder(szImportPath).Files
        Dim objFileName As String: objFileName = objFile.Name
        If (objFSO.GetExtensionName(objFileName) = "cls") Or _
            (objFSO.GetExtensionName(objFileName) = "frm") Or _
            (objFSO.GetExtensionName(objFileName) = "bas") Then
                Dim moduleName As String
                moduleName = Split(objFileName, ".")(0)
                If isWhiteListed(moduleName, whiteList) Then
                    tryRemoveModuleByName (moduleName)
                    On Error GoTo Problem_Importing
                    cmpComponents.Import objFile.Path
                    On Error GoTo 0
                    Debug.Print "Imported " & objFileName & " to Workbook"
                Else
                    Debug.Print moduleName & " not on white list. Skipped import"
                End If
        Else
            Debug.Print "Unrecognised extension. Skipped import for file: " & objFileName
        End If
    Next
    Debug.Print "***** Completed import from: " & importFolder
    Exit Sub
Problem_Importing:
    raiseErrorSync ("There was a problem importing the latest module.")
End Sub

Private Sub tryRemoveModuleByName(moduleName As String)
    Dim vbProj As VBIDE.VBProject
    Set vbProj = ThisWorkbook.VBProject
    With vbProj.VBComponents
        On Error GoTo Did_Not_Remove
        .Remove .Item(moduleName)
        On Error GoTo 0
        Debug.Print Now & " Removed module " & moduleName & " from " & ThisWorkbook.Name
    End With
    Exit Sub
Did_Not_Remove:
    Debug.Print Now & " Didn't remove module " & moduleName & " - Not Found."
End Sub

Private Sub testCreateFolderWithVBAFiles()
    MsgBox (createFolderWithVBAMiscFiles())
    MsgBox (createFolderWithProjectSpecificVBAFiles())
End Sub

Private Function createFolderWithProjectSpecificVBAFiles() As String
    Dim FSO As New FileSystemObject
    Dim totalPath As String: totalPath = getFolderWithProjectSpecificVbaFiles(FSO)
    If totalPath = "" Then totalPath = getWorkingDirPath(ThisWorkbook) & "ProjectSpecific"
    createFolderWithProjectSpecificVBAFiles = createFolderWithVBAFiles(FSO, totalPath)
End Function

Private Function createFolderWithVBAMiscFiles() As String
    Dim FSO As New FileSystemObject
    Dim totalPath As String: totalPath = getFolderWithVbaMiscFiles(FSO)
    If totalPath = "" Then totalPath = getWorkingDirPath(ThisWorkbook) & "GeneralPurpose"
    createFolderWithVBAMiscFiles = createFolderWithVBAFiles(FSO, totalPath)
End Function

Private Function createFolderWithVBAFiles(FSO As FileSystemObject, totalPath As String) As String
    
    If Not FSO.FolderExists(totalPath) Then
        On Error Resume Next
        Debug.Print "No folder exists. Attempting to make: " & totalPath
        MkDir totalPath
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(totalPath) = True Then
        createFolderWithVBAFiles = totalPath
    Else
        createFolderWithVBAFiles = "Error"
    End If
    
End Function

Public Function getWorkingDirPath(wb As Workbook)
    Dim prefixPath As String
    prefixPath = wb.Path
    If right(prefixPath, 1) <> "\" Then
        prefixPath = prefixPath & "\"
    End If
    getWorkingDirPath = prefixPath
End Function

Private Sub testGetSyncPropertyValue()
    MsgBox (getFolderWithVbaMiscFiles(New FileSystemObject))
    MsgBox (getFolderWithProjectSpecificVbaFiles(New FileSystemObject))
End Sub

Private Function getFolderWithProjectSpecificVbaFiles(FSO As FileSystemObject) As String
    getFolderWithProjectSpecificVbaFiles = getSyncPropertyValue(FSO, SPECIFIC_REL_KEY, SPECIFIC_ABS_KEY)
End Function

Private Function getFolderWithVbaMiscFiles(FSO As FileSystemObject) As String
    getFolderWithVbaMiscFiles = getSyncPropertyValue(FSO, MISC_REL_KEY, MISC_ABS_KEY)
End Function

'Relative trumps absolute
Private Function getSyncPropertyValue(FSO As FileSystemObject, relKey As String, absKey As String) As String
    If (FSO.FileExists(getConfigFileFullPath())) Then
        Dim textStream As textStream: Set textStream = getConfigInputStream(FSO)
        Do While (Not textStream.AtEndOfLine)
            Dim currLine As String: currLine = textStream.ReadLine
            Dim val As String
            If InStr(currLine, relKey) = 1 Then
                val = Replace(currLine, relKey, "", 1, 1)
                getSyncPropertyValue = getWorkingDirPath(ThisWorkbook) & val
                Exit Function
            ElseIf InStr(currLine, absKey) = 1 Then
                val = Replace(currLine, absKey, "", 1, 1)
                getSyncPropertyValue = val
                Exit Function
            End If
        Loop
    End If
    getSyncPropertyValue = ""
End Function

Private Sub testGetGenSyncListFullName()
    MsgBox (getGenSyncListFullName(New FileSystemObject))
End Sub

'Tries to get file name from main config file. Failing that, gives a default file name in current dir.
Private Function getGenSyncListFullName(FSO As FileSystemObject) As String
    Dim propertyName As String
    propertyName = getSyncPropertyValue(FSO, MISC_SYNCLIST_REL_KEY, MISC_SYNCLIST_ABS_KEY)
    If propertyName = "" Then propertyName = getWorkingDirPath(ThisWorkbook) & MISC_SYNCLIST_FILE_NAME
    getGenSyncListFullName = propertyName
End Function

Private Sub testGetInputStream()
    MsgBox ("Config file empty: " & getConfigInputStream(New FileSystemObject).AtEndOfStream)
    MsgBox ("Whitelist file empty: " & getSpecificWhitelistInputStream(New FileSystemObject).AtEndOfStream)
End Sub

Private Function getConfigInputStream(FSO As FileSystemObject) As textStream
    Dim fullPath As String: fullPath = getConfigFileFullPath()
    Set getConfigInputStream = getInputStream(FSO, fullPath)
End Function

Private Function getConfigFileFullPath() As String
    getConfigFileFullPath = getWorkingDirPath(ThisWorkbook) & CONFIG_FILE_NAME
End Function

Private Sub testIsWhiteListed()
    Dim moduleName1 As String, moduleName2 As String
    moduleName1 = "Dependencies"
    moduleName2 = "MyModule"
    MsgBox (moduleName1 & " is whitelisted: " & isWhiteListed(moduleName1, miscWhiteList()))
    MsgBox (moduleName2 & " is whitelisted: " & isWhiteListed(moduleName2, miscWhiteList()))
End Sub

Private Function isWhiteListed(moduleName As String, whiteList() As String) As Boolean
    isWhiteListed = mStringInArray(moduleName, whiteList)
End Function

Public Function mStringInArray(str As String, strArr() As String) As Boolean
    Dim strLooper As Variant
    On Error GoTo gtStrInArrErr
        For Each strLooper In strArr
            If str = strLooper Then
                mStringInArray = True
                Exit Function
            End If
        Next
    On Error GoTo 0
    mStringInArray = False
    Exit Function
gtStrInArrErr:
    Dim lErrorMsg As String
    lErrorMsg = "mStringInArray threw an error. Caused by: " & vbCrLf & _
        Err.Description
    raiseError (lErrorMsg)
End Function

Public Sub raiseError(msg As String)
    Err.Raise Number:=513, Description:=msg
End Sub

Private Sub testMiscWhiteList()
    Call MsgBox("General white list: " & vbCrLf & Join(miscWhiteList(), ","))
End Sub

Private Function miscWhiteList() As String()
    Dim FSO As New FileSystemObject
    Dim miscTs As textStream: Set miscTs = getMiscWhitelistInputStream(FSO)
    If miscTs.AtEndOfStream Then
        'Return an empty array here
        Dim arrayToReturn(0) As String
        miscWhiteList = arrayToReturn
        Exit Function
    End If
    miscWhiteList = Split(miscTs.ReadAll, vbNewLine)
End Function

Function getMiscWhitelistInputStream(FSO As FileSystemObject) As textStream
    Dim fullPath As String: fullPath = getGenSyncListFullName(FSO)
    Set getMiscWhitelistInputStream = getInputStream(FSO, fullPath)
End Function

Private Sub testSpecificWhiteList()
    Call MsgBox("Specific white list: " & vbCrLf & Join(specificWhiteList(), ","))
End Sub

Private Function specificWhiteList() As String()
    Dim FSO As New FileSystemObject
    Dim specTs As textStream: Set specTs = getSpecificWhitelistInputStream(FSO)
    If specTs.AtEndOfStream Then
        'Return an empty array here
        Dim arrayToReturn(0) As String
        specificWhiteList = arrayToReturn
        Exit Function
    End If
    specificWhiteList = Split(specTs.ReadAll, vbNewLine)
End Function

Function getSpecificWhitelistInputStream(FSO As FileSystemObject) As textStream
    Dim fullPath As String: fullPath = getSpecificWhitelistFileFullPath()
    Set getSpecificWhitelistInputStream = getInputStream(FSO, fullPath)
End Function

Private Function getSpecificWhitelistFileFullPath() As String
    getSpecificWhitelistFileFullPath = getWorkingDirPath(ThisWorkbook) & SPECIFIC_SYNCLIST_FILE_NAME
End Function

Function getInputStream(FSO As FileSystemObject, fullPath As String) As textStream
    On Error GoTo config_textStream_Error
    Set getInputStream = FSO.OpenTextFile(fullPath)
    On Error GoTo 0
    Exit Function
config_textStream_Error:
    raiseErrorSync ("Failed to connect text stream to file: " & fullPath)
End Function

Sub testgetWorkingDirPath()
    MsgBox (getWorkingDirPath(ThisWorkbook))
End Sub

Private Sub raiseErrorSync(msg As String)
    Err.Raise Number:=513, Description:=msg
End Sub

''DON'T ADD ANY TESTS TO SYNC - YOU'LL BREAK THE IMPORT
