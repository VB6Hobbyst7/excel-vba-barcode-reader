VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} log 
   Caption         =   "UserForm1"
   ClientHeight    =   4365
   ClientLeft      =   240
   ClientTop       =   945
   ClientWidth     =   9810.001
   OleObjectBlob   =   "log.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

Dim iRow As Long
Dim ws As Worksheet
Set ws = Worksheets("database")

'finds lowest row
iRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row

'input values
ws.Cells(iRow, 1).Value = Now
ws.Cells(iRow, 2).Value = Me.txtBarcode.Value

'clears field for next use
Me.txtBarcode = ""

Unload Me

End Sub

Private Sub cmdClose_Click()
Unload Me

End Sub

Private Sub UserForm_Click()

End Sub
