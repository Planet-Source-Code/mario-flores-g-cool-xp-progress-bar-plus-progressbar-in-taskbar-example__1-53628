VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Mario Flores Cool Xp ProgressBar"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8805
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   593
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   587
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImgList 
      Left            =   7440
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "More"
      Height          =   375
      Left            =   6720
      TabIndex        =   17
      Top             =   7080
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Search Style Scrolling"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   16
      Top             =   5760
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Smooth Scrolling"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   5280
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Standard Scrolling"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   4800
      Value           =   -1  'True
      Width           =   1935
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   8400
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   4604
            Text            =   "Mario Flores G Cool Xp ProgressBar"
            TextSave        =   "Mario Flores G Cool Xp ProgressBar"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "5/8/2004"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6840
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3015
      Left            =   2760
      MouseIcon       =   "Form1.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":06DC
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   3
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Use Percent Text"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   7800
      Width           =   1815
   End
   Begin Project1.XP_ProgressBar XP_ProgressBar1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2760
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin Project1.XP_ProgressBar XP_ProgressBarx 
      Height          =   255
      Left            =   5880
      TabIndex        =   1
      Top             =   6240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
   End
   Begin Project1.XP_ProgressBar XP_ProgressBarSilver 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin Project1.XP_ProgressBar XP_ProgressBarOlive 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin Project1.XP_ProgressBar XP_ProgressBarBlue 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin Project1.XP_ProgressBar XP_ProgressBarTask 
      Height          =   480
      Left            =   7080
      TabIndex        =   18
      Top             =   4680
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "This Progress Is used To display The Form Icon Progress"
      Height          =   555
      Left            =   6120
      TabIndex        =   19
      Top             =   5400
      Width           =   2385
   End
   Begin VB.Label LblValue 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   3960
      TabIndex        =   12
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Color:"
      Height          =   195
      Left            =   2880
      TabIndex        =   11
      Top             =   7200
      Width           =   405
   End
   Begin VB.Shape ShapeColor 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3480
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Example Color Blue XP"
      Height          =   195
      Left            =   1920
      TabIndex        =   10
      Top             =   240
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Example Color Olive XP"
      Height          =   195
      Left            =   1920
      TabIndex        =   9
      Top             =   960
      Width           =   1665
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Example Color Silver XP"
      Height          =   195
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Picker Color Example"
      Height          =   195
      Left            =   3480
      TabIndex        =   7
      Top             =   2520
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Dim MTime  As Long
Dim MDown  As Boolean

Const XPBlue_ProgressBar = &H2BD228
Const XPOlive_ProgressBar = &H4A86E4
Const XPSilver_ProgressBar = &H76AE83



Private Sub Check2_Click()
XP_ProgressBar1.ShowText = IIf(Check2.Value = 0, False, True)
XP_ProgressBarBlue.ShowText = IIf(Check2.Value = 0, False, True)
XP_ProgressBarOlive.ShowText = IIf(Check2.Value = 0, False, True)
XP_ProgressBarSilver.ShowText = IIf(Check2.Value = 0, False, True)
XP_ProgressBarx.ShowText = IIf(Check2.Value = 0, False, True)
XP_ProgressBarTask.ShowText = IIf(Check2.Value = 0, False, True)
End Sub

Private Sub Command1_Click()
Form2.Show
End Sub


Private Sub Form_Initialize()
InitCommonControls

'******** Set Form's ShowinTaskBar Property = False ....
'If you Want to Disable Window Showing the ProgressBar in Running Applications ToolBar...
'Note: Windows doesn't refresh this Toolbar in real time so
'the Percent and Image here are never going to be accurate ..

'But not my fault….


MsgBox "Watch The Form Title Bar And The TaskBar …", vbInformation, "See This"

End Sub


Private Sub Form_Load()

XP_ProgressBar1.Max = 100
XP_ProgressBar1.Min = 1

XP_ProgressBarBlue.Max = 100
XP_ProgressBarBlue.Min = 1

XP_ProgressBarOlive.Max = 100
XP_ProgressBarOlive.Min = 1

XP_ProgressBarSilver.Max = 100
XP_ProgressBarSilver.Min = 1

XP_ProgressBarx.Max = 100
XP_ProgressBarx.Min = 1

XP_ProgressBarTask.Max = 100
XP_ProgressBarTask.Min = 1


XP_ProgressBarBlue.Color = XPBlue_ProgressBar
XP_ProgressBarOlive.Color = XPOlive_ProgressBar
XP_ProgressBarSilver.Color = XPSilver_ProgressBar

XP_ProgressBar1.Color = vbHighlight
XP_ProgressBarx.Color = vbHighlight
XP_ProgressBarTask.Color = vbHighlight

ShapeColor.BackColor = vbHighlight

ShowProgressInStatusBar XP_ProgressBarx, StatusBar1, 3

XP_ProgressBarTask.ShowInTask = True ''FALSE VALUE DISABLES PROGRESS IN TASKBAR


End Sub

Private Sub Form_Unload(Cancel As Integer) 'on form unload
    RemoveTray
End Sub

Private Sub Option1_Click(Index As Integer)

If Index = 2 Then
    Timer1.Interval = 20
Else                         'Time Interval yust to Show Search Style Demo.
    Timer1.Interval = 100
End If

XP_ProgressBar1.Scrolling = Index
XP_ProgressBarBlue.Scrolling = Index
XP_ProgressBarOlive.Scrolling = Index
XP_ProgressBarSilver.Scrolling = Index
XP_ProgressBarx.Scrolling = Index
XP_ProgressBarTask.Scrolling = Index

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim R      As Integer
Dim G      As Integer
Dim B      As Integer
Dim PixCol As Long

PixCol = GetPixel(Picture1.HDC, X, Y)

'Convert to RGB
R = PixCol Mod 256
B = Int(PixCol / 65536)
G = (PixCol - (B * 65536) - R) / 256

If R < 0 Then R = 0
If G < 0 Then G = 0
If B < 0 Then B = 0


ShapeColor.BackColor = RGB(R, G, B)
XP_ProgressBar1.Color = ShapeColor.BackColor
XP_ProgressBarx.Color = ShapeColor.BackColor
XP_ProgressBarTask.Color = ShapeColor.BackColor

MDown = True


End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MDown Then Picture1_MouseDown Button, Shift, X, Y
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MDown = False

End Sub

Private Sub Timer1_Timer()
MTime = MTime + 1

If MTime > XP_ProgressBar1.Max Then
    MTime = XP_ProgressBar1.Min
End If



XP_ProgressBar1.Value = MTime
XP_ProgressBarBlue.Value = MTime
XP_ProgressBarOlive.Value = MTime
XP_ProgressBarSilver.Value = MTime
XP_ProgressBarx.Value = MTime
XP_ProgressBarTask.Value = MTime

LblValue = (XP_ProgressBar1.Max * XP_ProgressBar1.Value) / 100 & " %"


Me.Caption = LblValue


Call RefreshTray(Me, LblValue)



End Sub




