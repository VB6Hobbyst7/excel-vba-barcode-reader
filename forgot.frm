VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} forgot 
   Caption         =   "UserForm1"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6645
   OleObjectBlob   =   "forgot.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "forgot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseButton_Click()
Unload Me
End Sub

Private Sub SubmitButton_Click()

Dim ws As Worksheet
Set ws = Worksheets("database")

'phone number length check
If Len(Me.txtPhone.Value) = 11 Then
    
    'handler for non-existing phone number
    On Error GoTo ErrHandler
    
    
    
    
    ''''''
    'uses field txtPhone to find an entry matching
    'selects cell 3 spaces to the left relative to first entry
    'saves barcode value at N4
    '
    'uses existing barcode value as input to a normal data entry
    '
    ''''''
    
    Cells.Find(What:=Me.txtPhone.Value, After:=ActiveCell, LookIn:=xlFormulas2, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, -3).Range("A1").Select
    Selection.Copy
    Range("M4").Select
    ActiveSheet.Paste
    
    'finds lowest row
    iRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row

    ws.Cells(iRow, 1).Value = Now 'timestamp
    ws.Cells(iRow, 2).Value = Range("M4").Value 'forgotten barcode value
    
    'cleans for later use
    Me.txtPhone.Value = ""
    Unload Me
Else
    MsgBox ("Incomplete phone number")
    Unload Me

End If
Exit Sub

'non-present phone number handler
ErrHandler:
    Me.txtPhone.Value = ""
MsgBox ("Phone number not found." & vbCrLf & _
         "Please register new person.")
End Sub

Private Sub UserForm_Click()

End Sub
