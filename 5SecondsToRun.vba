Public strFileName As String
Public strFileName1 As String
Public columnTouse As Excel.Range
Public columnTouse1 As Integer
Public columnTouse2 As Excel.Range
Public columnTouse3 As Integer
Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

Sub OptimizeCode_Begin()

Application.ScreenUpdating = False

EventState = Application.EnableEvents
Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False

End Sub
Sub OptimizeCode_End()

ActiveSheet.DisplayPageBreaks = PageBreakState
Application.Calculation = CalcState
Application.EnableEvents = EventState
Application.ScreenUpdating = True

End Sub
Sub test()
    



End Sub

Private Sub CommandButton1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
End Sub

Private Sub CommandButton1_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)

End Sub

Private Sub CommandButton2_Click()

End Sub

Public Sub GetChildPath_Click()
With Application.FileDialog(msoFileDialogFilePicker)
    .Show
    strFileName1 = .SelectedItems(1)
End With
End Sub
Public Sub GetChildPath()
With Application.FileDialog(msoFileDialogFilePicker)
    .Show
    strFileName1 = .SelectedItems(1)
End With
End Sub

Private Sub GetChildPath_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub GetMsterPath_Click()

End Sub

Public Sub GetMsterPath_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
With Application.FileDialog(msoFileDialogFilePicker)
    .Show
    strFileName = .SelectedItems(1)
End With
End Sub

Public Sub getmasterpath()
With Application.FileDialog(msoFileDialogFilePicker)
    .Show
    strFileName = .SelectedItems(1)
End With
End Sub
Private Sub SyncData_Click()

End Sub
Public Sub SyncData_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Public Sub columnusage(namesearch As Variant, ws As Worksheet)
Dim rc As Excel.Range
Set rc = ws.Range("A1:Z80").Find(what:=namesearch, lookat:=xlWhole)
If Not rc Is Nothing Then
Set columnTouse = rc
Exit Sub
End If
End Sub
Public Sub columnusage1(namesearch As Variant, ws As Worksheet)
Dim rc1 As Excel.Range
Set rc1 = ws.Range("A1:Z80").Find(what:=namesearch, lookat:=xlWhole)
If Not rc1 Is Nothing Then
columnTouse1 = rc1.Column
Exit Sub
End If
End Sub
Public Sub columnusage2(namesearch As Variant, ws As Worksheet)
Dim rc2 As Excel.Range
Set rc2 = ws.Range("A1:Z80").Find(what:=namesearch, lookat:=xlWhole)
If Not rc2 Is Nothing Then
Set columnTouse2 = rc2
Exit Sub
End If
End Sub

Public Sub columnusage3(namesearch As Variant, ws As Worksheet)
Dim rc3 As Excel.Range
Set rc3 = ws.Range("A1:Z80").Find(what:=namesearch, lookat:=xlWhole)
If Not rc3 Is Nothing Then
columnTouse3 = rc3.Column
Exit Sub
End If
End Sub

Sub WaitFor(NumOfSeconds As Long)
Dim SngSec As Long
SngSec = Timer + NumOfSeconds

Do While Timer < SngSec
DoEvents
Loop

End Sub

Public Sub syncdata()
Dim wbkA As Workbook
    Dim wbkB As Workbook
    Dim i As Integer, a As Integer, k As String, b As Integer
    Dim j As Excel.Range
    Dim columnNamesRow As Integer, lastUsedColumn As Integer, col2 As Integer, col3 As Integer
    Dim columnNamesRow1 As Integer, lastUsedColumn1 As Integer
    Dim nameToSearch As String, nameToSearch1 As String
    Dim maxrows As Long, maxrows1 As Long, abc As String, sht As Worksheet
    Dim pivot As Variant, destination As Variant
    
Call OptimizeCode_Begin

MsgBox "Select source file"
Call getmasterpath
If strFileName <> "" Then
MsgBox "Select destination file"
Call GetChildPath
End If
    
Set wbkA = Workbooks.Open(Filename:=strFileName)
    
Set wbkB = Workbooks.Open(Filename:=strFileName1)

pivot = InputBox("Enter pivot column title")
If pivot <> "" Then
destination = InputBox("Enter destination column name")
End If

For b = 1 To wbkB.Worksheets.Count
Call columnusage2(pivot, wbkB.Sheets(b))
Call columnusage3(destination, wbkB.Sheets(b))

For a = 1 To wbkA.Worksheets.Count
Call columnusage(pivot, wbkA.Sheets(a))
Call columnusage1(destination, wbkA.Sheets(a))


If Not columnTouse Is Nothing And columnTouse1 <> 0 And Not columnTouse2 Is Nothing And columnTouse3 <> 0 Then

 For i = 2 To 250
 If Not IsEmpty(wbkB.Sheets(b).Cells(i, columnTouse2.Column).Value) Then
        k = wbkB.Sheets(b).Cells(i, columnTouse2.Column).Value
        If k <> "" Then
        Set j = wbkA.Sheets(a).Range("A1:Z250").Find(what:=k, lookat:=xlWhole)
        Set sht = wbkA.Sheets(a)
        If Not j Is Nothing Then
        abc = wbkA.Sheets(a).Cells(j.Row, columnTouse1).Value
        If abc <> "" Then
        wbkB.Sheets(b).Cells(i, columnTouse3).Value = abc
        End If
        End If
        End If
        End If
        Next i

End If


Next a
Next b

wbkA.Close savechanges:=True
wbkB.Close savechanges:=True

MsgBox "Sync complete"

Call OptimizeCode_End

End Sub
