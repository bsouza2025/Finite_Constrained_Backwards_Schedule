Attribute VB_Name = "Module4"
' Module: ExportAllVBA_Safe
Option Explicit

' VBIDE types via late binding so no reference required.
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

Public Sub ExportAllModules_Safe()
    On Error GoTo Fail

    ' 0) Require trust to VBA project object model
    If Not IsVBATrusted() Then
        MsgBox "Enable 'Trust access to the VBA project object model' first:" & vbCrLf & _
               "File ? Options ? Trust Center ? Trust Center Settings ? Macro Settings.", vbExclamation
        Exit Sub
    End If

    ' 1) Pick an export folder (avoids empty ThisWorkbook.Path problem)
    Dim exportPath As String
    exportPath = PickFolder()
    If Len(exportPath) = 0 Then Exit Sub ' user cancelled

    ' Ensure trailing slash
    If Right$(exportPath, 1) <> "\" Then exportPath = exportPath & "\"

    ' 2) Create folder if missing
    CreateFolderIfMissing exportPath

    ' 3) Export each component
    Dim vbComp As Object    ' VBIDE.VBComponent (late bound)
    Dim filePath As String, baseName As String, ext As String

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule
                ext = ".bas"
            Case vbext_ct_ClassModule
                ext = ".cls"
            Case vbext_ct_MSForm
                ext = ".frm"
            Case vbext_ct_Document
                ' Skip Sheet/Workbook code-behind by default. Uncomment to export:
                ' ext = ".cls"
            Case Else
                ext = ""
        End Select

        If Len(ext) > 0 Then
            baseName = SanitizeFileName(CStr(vbComp.Name))
            filePath = UniquePath(exportPath & baseName & ext)
            vbComp.Export filePath
        End If
    Next vbComp

    MsgBox "Export complete to:" & vbCrLf & exportPath, vbInformation
    Exit Sub

Fail:
    MsgBox "Export failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ---------- helpers ----------

Private Function IsVBATrusted() As Boolean
    ' If access is blocked, reading VBProject.Name throws 1004 in most builds.
    On Error GoTo NotTrusted
    Dim tmp As String
    tmp = ThisWorkbook.VBProject.Name
    IsVBATrusted = True
    Exit Function
NotTrusted:
    IsVBATrusted = False
End Function

Private Function PickFolder() As String
    ' Uses FileDialog to avoid bad paths
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select folder to export VBA modules"
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        Else
            PickFolder = vbNullString
        End If
    End With
End Function

Private Sub CreateFolderIfMissing(ByVal path As String)
    If Len(path) = 0 Then Exit Sub
    If Dir(path, vbDirectory) = "" Then MkDir path
End Sub

Private Function SanitizeFileName(ByVal nameIn As String) As String
    ' Remove characters invalid in Windows filenames:  \ / : * ? " < > |
    Dim badChars As Variant, c As Variant, s As String
    s = nameIn
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each c In badChars
        s = Replace$(s, CStr(c), "_")
    Next
    If Len(Trim$(s)) = 0 Then s = "Module"
    SanitizeFileName = s
End Function

Private Function UniquePath(ByVal fullPath As String) As String
    ' If file exists, append (1), (2), ...
    Dim base As String, ext As String, i As Long
    i = InStrRev(fullPath, ".")
    If i > 0 Then
        base = Left$(fullPath, i - 1)
        ext = Mid$(fullPath, i)
    Else
        base = fullPath
        ext = ""
    End If

    Dim attempt As String, n As Long
    attempt = fullPath
    n = 1
    Do While Len(Dir$(attempt, vbNormal)) > 0
        attempt = base & " (" & n & ")" & ext
        n = n + 1
    Loop
    UniquePath = attempt
End Function




