VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uncoop 
   Caption         =   "UserForm1"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5940
   OleObjectBlob   =   "uncoop.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uncoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseButton_Click()
Unload Me
End Sub



Private Sub SubmitButton_Click()


Dim iRow As Long
Dim ws As Worksheet
Set ws = Worksheets("database")

'correct length of phone number check block
If Len(Me.txtPhone.Value) = 11 Then

'finding the lowest row in the sheet
iRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row

'value input
ws.Cells(iRow, 1).Value = Now
ws.Cells(iRow, 3).Value = Me.txtNameF.Value
ws.Cells(iRow, 4).Value = Me.txtNameS.Value
ws.Cells(iRow, 5).Value = Me.txtPhone.Value

'cleans userform for next entry
Me.txtNameF.Value = ""
Me.txtNameS.Value = ""
Me.txtPhone.Value = ""

Unload Me
Else
    MsgBox ("Incomplete phone number")

End If
End Sub
