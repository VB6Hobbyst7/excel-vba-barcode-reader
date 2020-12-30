VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} registerNew 
   Caption         =   "Register"
   ClientHeight    =   9375.001
   ClientLeft      =   300
   ClientTop       =   1170
   ClientWidth     =   10860
   OleObjectBlob   =   "registerNew.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "registerNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

Dim iRow As Long
Dim ws As Worksheet
Set ws = Worksheets("database")

'check if phone number is correct length
If Len(Me.txtPhone.Value) = 11 Then

'find lowest row
iRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row

'value entry
ws.Cells(iRow, 1).Value = Now 'timestamp
ws.Cells(iRow, 2).Value = Me.txtBarcode 'barcode
ws.Cells(iRow, 3).Value = Me.txtNameF.Value 'first name
ws.Cells(iRow, 4).Value = Me.txtNameS.Value 'last name
ws.Cells(iRow, 5).Value = Me.txtPhone.Value 'phone number
ws.Cells(iRow, 6).Value = Me.txtPostcode.Value 'postcode
ws.Cells(iRow, 7).Value = Me.txtA.Value 'how many adults on premises
ws.Cells(iRow, 8).Value = Me.txtC.Value 'how many children on premises
ws.Cells(iRow, 9).Value = Me.txtOc.Value 'occupation
ws.Cells(iRow, 10).Value = Me.togY.Value 'housing association boolean

'clean fields for next entry
Me.txtBarcode = ""
Me.txtNameF.Value = ""
Me.txtNameS.Value = ""
Me.txtPostcode.Value = ""
Me.txtPhone.Value = ""
Me.txtA.Value = ""
Me.txtC.Value = ""
Me.txtOc.Value = ""
Me.togY.Value = ""
Me.togN.Value = ""

Unload Me
Else
    MsgBox ("Incomplete phone number")

End If

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
