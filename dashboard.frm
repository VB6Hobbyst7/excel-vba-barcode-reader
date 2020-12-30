VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dashboard 
   Caption         =   "UserForm1"
   ClientHeight    =   11670
   ClientLeft      =   420
   ClientTop       =   1590
   ClientWidth     =   21210
   OleObjectBlob   =   "dashboard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdHelp_Click()
problem.Show
End Sub


Private Sub ForgotButton_Click()
forgot.Show
End Sub

Private Sub RegisterButton_Click()
registerNew.Show
End Sub

Private Sub RudeButton_Click()
uncoop.Show
End Sub


Private Sub ScanButton_Click()
log.Show
End Sub

Private Sub CloseButton_Click()
ActiveWorkbook.Save
Unload Me
End Sub

Private Sub LostButton_Click()
lost.Show
End Sub

Private Sub dashboard_Initialize()
With Application
.WindowState = xlMaximized
Zoom = Int(.Width / Me.Width * 100)
Width = .Width
Height = .Height
End With
End Sub

Private Sub UserForm_Click()

End Sub
