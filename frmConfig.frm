VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "Lake Screensaver Configuration"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6465
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBeach 
      Caption         =   "Show Beach"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3200
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkBackground 
      Caption         =   "Show Background"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2700
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkShip 
      Caption         =   "Show ship animation"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2200
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CheckBox chkFish 
      Caption         =   "Show fish animation"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1700
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CheckBox chkBird 
      Caption         =   "Show bird animation"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.TextBox txtOpacity 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Text            =   "200"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Lake Screensaver              Mike Meiskey Lake@etherealstorm.com Copyright 2005"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3000
      TabIndex        =   7
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Water Transparency (0=Transparent 255=Opaque)"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    bCancel = True
    DoEvents
    UnloadAll
    'End
End Sub

Private Sub cmdOK_Click()
    ' save stuff
    bCancel = True
    WriteIni
    UnloadAll
    'End
End Sub

Private Sub Form_Load()
    SetControlValues
End Sub

Private Sub SetControlValues()
    
    If birdAnim Then
        chkBird.Value = vbChecked
    Else
        chkBird.Value = vbUnchecked
    End If
    
    If fishAnim Then
        chkFish.Value = vbChecked
    Else
        chkFish.Value = vbUnchecked
    End If
    
    If shipAnim Then
        chkShip.Value = vbChecked
    Else
        chkShip.Value = vbUnchecked
    End If
    
    If showBack Then
        chkBackground.Value = vbChecked
    Else
        chkBackground.Value = vbUnchecked
    End If
    
    If showBeach Then
        chkBeach.Value = vbChecked
    Else
        chkBeach.Value = vbUnchecked
    End If
    
    txtOpacity.Text = waterTransparency
    
End Sub

Private Sub txtOpacity_Change()
    If Val(txtOpacity.Text) < 0 Or Val(txtOpacity.Text) > 255 Then
        MsgBox "Value must be between 0 and 255"
    End If
End Sub

Private Sub WriteIni()
    
    Dim iniPath As String
    Dim fFile As Integer
    Dim iniString As String
    
    fFile = FreeFile
    
    If chkBird.Value = vbChecked Then
        birdAnim = True
    Else
        birdAnim = False
    End If
    
    If chkFish.Value = vbChecked Then
        fishAnim = True
    Else
        fishAnim = False
    End If
    
    If chkShip.Value = vbChecked Then
        shipAnim = True
    Else
        shipAnim = False
    End If
    
    If chkBackground.Value = vbChecked Then
        showBack = True
    Else
        showBack = False
    End If
    
    If chkBeach.Value = vbChecked Then
        showBeach = True
    Else
        showBeach = False
    End If
    
    waterTransparency = Val(txtOpacity.Text)
    
    iniString = "birdAnim=" & birdAnim & vbCrLf & "fishAnim=" & fishAnim & vbCrLf & "shipAnim=" & shipAnim & vbCrLf & "waterTransparency=" & waterTransparency & vbCrLf & "showback=" & showBack & vbCrLf & "showBeach=" & showBeach & vbCrLf
    
    If Right(App.Path, 1) <> "\" Then
        iniPath = App.Path & "\LakeScreenSaver.ini"
    Else
        iniPath = App.Path & "LakeScreenSaver.ini"
    End If
    
    Open iniPath For Output As #fFile
        Write #fFile, iniString
    Close #fFile

End Sub

