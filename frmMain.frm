VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MouseColor - Press F7 to Save Color"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSaved 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   16
      Top             =   480
      Width           =   375
   End
   Begin VB.Frame fmeRGB 
      Caption         =   "RGB and Hex Color Values"
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4095
      Begin VB.PictureBox picCurrent 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtHexS 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtHexC 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtBlueC 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtGreenC 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtRedC 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtBlueS 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtGreenS 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtRedS 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblHex 
         Caption         =   "Hex"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblBlue 
         Caption         =   "Blue"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblGreen 
         Caption         =   "Green"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblRed 
         Caption         =   "Red"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Caption         =   "Current"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblSaved 
         Caption         =   "Saved"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   -240
      Top             =   -240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 118 Then
        picSaved.BackColor = picCurrent.BackColor
        txtRedS.Text = txtRedC.Text
        txtGreenS.Text = txtGreenC.Text
        txtBlueS.Text = txtBlueC.Text
        txtHexS.Text = txtHexC.Text
        Clipboard.Clear
        Clipboard.SetText "(" & txtRedS.Text & "," & txtGreenS.Text & "," & txtBlueS.Text & ") " & txtHexS.Text
    End If
End Sub

Private Sub tmrUpdate_Timer()
    Dim Pt As POINTAPI
    DisplayHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    GetCursorPos Pt
    picCurrent.BackColor = GetPixel(DisplayHDC, Pt.x, Pt.y)
    txtRedC.Text = LongToRGB(GetPixel(DisplayHDC, Pt.x, Pt.y)).R
    txtGreenC.Text = LongToRGB(GetPixel(DisplayHDC, Pt.x, Pt.y)).G
    txtBlueC.Text = LongToRGB(GetPixel(DisplayHDC, Pt.x, Pt.y)).B
    txtHexC.Text = Hex(GetPixel(DisplayHDC, Pt.x, Pt.y))
End Sub
