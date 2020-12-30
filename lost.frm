VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} lost 
   ClientHeight    =   6270
   ClientLeft      =   300
   ClientTop       =   1185
   ClientWidth     =   9555.001
   OleObjectBlob   =   "lost.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "lost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheck_Click()

Dim ws As Worksheet
Set ws = Worksheets("database")


'phone number length checker
If Len(Me.txtPhone.Value) = 11 Then

    'incorrect phone number handler
    On Error GoTo ErrHandler
    
    
    ''''''
    'uses field txtPhone to find an entry matching
    'selects cell 3 spaces to the left relative to first entry
    'saves old barcode value at M4
    'replaces all old entries of occurances of the previous barcode with value at txtBarcode within the current userform
    '
    'reason for saving old barcode is the data gets overwritten before before it can be used
    '
    ''''''
    
    Sheets("database").Select
      
      Cells.Find(What:=Me.txtPhone.Value, After:=ActiveCell, LookIn:=xlFormulas2, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, -3).Range("Table3[[#Headers],[Date and Time]]").Select
    Selection.Copy
    Range("L4").Select
    ActiveSheet.Paste
    Cells.Replace What:=Range("L4").Value, Replacement:=Me.txtBarcode.Value, LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    'cleans fields for later use
    Me.txtBarcode.Value = ""
    Me.txtPhone.Value = ""
    
    Unload Me
Else
    MsgBox ("Incomplete phone number")
    Unload Me
    
End If
Exit Sub

'clears fields, preventing crash of program when phone number is not present in the database
ErrHandler:
    Me.txtBarcode.Value = ""
    Me.txtPhone.Value = ""
MsgBox ("Phone number not found")
End Sub


Private Sub cmdClose_Click()
Unload Me
End Sub


Private Sub UserForm_Click()

End Sub
