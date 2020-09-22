VERSION 5.00
Begin VB.Form frmPreview 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Preview"
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   LinkTopic       =   "Form1"
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   30
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
    SetStretchBltMode Me.hdc, vbPaletteModeNone
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 1, ByVal 0&, 0
    bCancel = True
'    DoEvents
    UnloadAll
'    End
End Sub
