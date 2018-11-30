VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmSettings 
   Caption         =   "Email Generator Settings"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5910
   OleObjectBlob   =   "FrmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnSave_Click()
    ChkSave.Value = True
    FrmSettings.Hide
End Sub

Private Sub BtnCancel_Click()
    ChkSave.Value = False
    FrmSettings.Hide
End Sub

Private Sub UserForm_Activate()
    MsgBox Prompt:="TODO: Load settings."
End Sub


