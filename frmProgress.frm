VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "sending progress..."
   ClientHeight    =   1404
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8784.001
   OleObjectBlob   =   "frmProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClose_Click()
    frmProgress.Hide
    Set frmProgress = Nothing
End Sub
