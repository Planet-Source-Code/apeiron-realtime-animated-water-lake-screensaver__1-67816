VERSION 5.00
Begin VB.Form frmDisplay 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   LinkTopic       =   "Form2"
   ScaleHeight     =   4395
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XPos As Single
Dim YPos As Single

Private Sub Form_Load()
    'SetStretchBltMode Me.hdc, vbPaletteModeNone
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Abs(XPos - X) > 5 And Abs(YPos - Y) > 5 And XPos <> 0 And YPos <> 0 Then
        Unload Me
        bCancel = True
    Else
        XPos = X
        YPos = Y
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bCancel = True
    ShowCursor True
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 1, ByVal 0&, 0
End Sub
