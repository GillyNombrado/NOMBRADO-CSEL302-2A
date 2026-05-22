VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Student Form"
   ClientHeight    =   6060
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   12210
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ws As Worksheet
Dim Gender As String
Dim Subjects As String
Dim lastRow As Long
Dim r As Range

Private Sub SetWorksheet()
    Set ws = ThisWorkbook.Sheets("Sheet1")
End Sub

Private Function GetSelectedSubjects() As String
    Dim selected As String
    selected = ""
    If CheckBox1.Value = True Then selected = selected & CheckBox1.Caption & " "
    If CheckBox2.Value = True Then selected = selected & CheckBox2.Caption & " "
    If CheckBox3.Value = True Then selected = selected & CheckBox3.Caption & " "
    GetSelectedSubjects = Trim(selected)
End Function

Private Sub SaveToSheet(rowNum As Long)
    Call SetWorksheet
    ws.Cells(rowNum, 1).Value = IDtxtb.Text
    ws.Cells(rowNum, 2).Value = FirstNametxtb.Text
    ws.Cells(rowNum, 3).Value = LastNametxtb.Text
    ws.Cells(rowNum, 4).Value = CourseCB.Text
    ws.Cells(rowNum, 5).Value = YearCB.Text
    ws.Cells(rowNum, 6).Value = Gender
    ws.Cells(rowNum, 7).Value = GetSelectedSubjects()
End Sub

Private Sub LoadFromSheet(rowNum As Long)
    Call SetWorksheet
    IDtxtb.Text = ws.Cells(rowNum, 1).Value
    FirstNametxtb.Text = ws.Cells(rowNum, 2).Value
    LastNametxtb.Text = ws.Cells(rowNum, 3).Value
    CourseCB.Text = ws.Cells(rowNum, 4).Value
    YearCB.Text = ws.Cells(rowNum, 5).Value
    
    Gender = ws.Cells(rowNum, 6).Value
    If Gender = "Female" Then
        OptionButton1.Value = True
    ElseIf Gender = "Male" Then
        OptionButton2.Value = True
    Else
        OptionButton1.Value = False
        OptionButton2.Value = False
    End If
    
    Dim dbSubjects As String
    dbSubjects = ws.Cells(rowNum, 7).Value
    CheckBox1.Value = (InStr(1, dbSubjects, CheckBox1.Caption, vbTextCompare) > 0)
    CheckBox2.Value = (InStr(1, dbSubjects, CheckBox2.Caption, vbTextCompare) > 0)
    CheckBox3.Value = (InStr(1, dbSubjects, CheckBox3.Caption, vbTextCompare) > 0)
End Sub

Private Sub AddBtn_Click()
    Call SetWorksheet
    
    If Trim(IDtxtb.Text) = "" Then
        MsgBox "Please enter a Student ID.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    Set r = ws.Columns(1).Find(IDtxtb.Text, LookIn:=xlValues, LookAt:=xlWhole)
    If Not r Is Nothing Then
        MsgBox "This Student ID already exists! Use the Update button to modify it.", vbExclamation, "Duplicate ID"
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Call SaveToSheet(lastRow)
    
    MsgBox "Record Added Successfully!", vbInformation, "Success"
    Call ClearBtn_Click ' Clear the form fields
End Sub

Private Sub SearchBtn_Click()
    Call SetWorksheet
    
    If Trim(IDtxtb.Text) = "" Then
        MsgBox "Please enter an ID to search.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    Set r = ws.Columns(1).Find(IDtxtb.Text, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not r Is Nothing Then
        Call LoadFromSheet(r.Row)
        MsgBox "Record Found!", vbInformation, "Search Result"
    Else
        MsgBox "Record Not Found.", vbCritical, "Search Result"
    End If
End Sub

Private Sub UpdateBtn_Click()
    Call SetWorksheet
    
    If Trim(IDtxtb.Text) = "" Then
        MsgBox "Please enter an ID to update.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    Set r = ws.Columns(1).Find(IDtxtb.Text, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not r Is Nothing Then
        Call SaveToSheet(r.Row)
        MsgBox "Record Updated!", vbInformation, "Success"
    Else
        MsgBox "ID not found. Cannot update.", vbCritical, "Error"
    End If
End Sub

Private Sub DeleteBtn_Click()
    Call SetWorksheet
    
    If Trim(IDtxtb.Text) = "" Then
        MsgBox "Please enter an ID to delete.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    Set r = ws.Columns(1).Find(IDtxtb.Text, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not r Is Nothing Then
        If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
            ws.Rows(r.Row).Delete
            Call CommandButton5_Click ' Clear form fields
            MsgBox "Record Deleted.", vbInformation, "Success"
        End If
    Else
        MsgBox "ID not found. Cannot delete.", vbCritical, "Error"
    End If
End Sub

Private Sub ClearBtn_Click()
    IDtxtb.Text = ""
    FirstNametxtb.Text = ""
    LastNametxtb.Text = ""
    CourseCB.ListIndex = -1
    YearCB.ListIndex = -1
    
    OptionButton1.Value = False
    OptionButton2.Value = False
    Gender = ""
    
    CheckBox1.Value = False
    CheckBox2.Value = False
    CheckBox3.Value = False
    Subjects = ""
End Sub

Private Sub OptionButton1_Click()
    If OptionButton1.Value = True Then Gender = "Female"
End Sub

Private Sub OptionButton2_Click()
    If OptionButton2.Value = True Then Gender = "Male"
End Sub

Private Sub UserForm_Initialize()
    CourseCB.AddItem "BSCS"
    CourseCB.AddItem "BSIT"
    
    YearCB.AddItem "1st"
    YearCB.AddItem "2nd"
    YearCB.AddItem "3rd"
    YearCB.AddItem "4th"
End Sub

Private Sub CheckBox1_Click()
    Subjects = GetSelectedSubjects()
End Sub

Private Sub CheckBox2_Click()
    Subjects = GetSelectedSubjects()
End Sub

Private Sub CheckBox3_Click()
    Subjects = GetSelectedSubjects()
End Sub
