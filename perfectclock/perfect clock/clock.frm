VERSION 5.00
Begin VB.Form Perfectclock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Perfect clock created dundee_united70@hotmail.com"
   ClientHeight    =   6555
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5145
   Icon            =   "clock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Height          =   15
      Left            =   2280
      Picture         =   "clock.frx":0442
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox Picture4 
      Height          =   15
      Left            =   4080
      Picture         =   "clock.frx":665F
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox Picture3 
      Height          =   15
      Left            =   2640
      Picture         =   "clock.frx":BF06
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox Picture2 
      Height          =   15
      Left            =   4080
      Picture         =   "clock.frx":11845
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   3120
      Picture         =   "clock.frx":16E10
      ScaleHeight     =   135
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1560
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   240
      Top             =   1200
   End
   Begin VB.Line hours 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   5
      X1              =   2760
      X2              =   3240
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line seconds 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   2760
      X2              =   3600
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line minutes 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      X1              =   2640
      X2              =   2640
      Y1              =   2040
      Y2              =   3120
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404080&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   6615
      Left            =   0
      Picture         =   "clock.frx":1BD95
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5175
   End
   Begin VB.Menu sel 
      Caption         =   "&Select Favorate "
      Index           =   1
      Begin VB.Menu r 
         Caption         =   "&Red"
         Index           =   1
         Shortcut        =   ^R
      End
      Begin VB.Menu v 
         Caption         =   "&Voilet"
         Index           =   2
         Shortcut        =   ^V
      End
      Begin VB.Menu bl 
         Caption         =   "&Blue"
         Index           =   3
         Shortcut        =   ^B
      End
      Begin VB.Menu gr 
         Caption         =   "&Green"
         Index           =   4
         Shortcut        =   ^G
      End
      Begin VB.Menu ye 
         Caption         =   "&Yellow"
         Index           =   5
         Shortcut        =   ^Y
      End
   End
End
Attribute VB_Name = "Perfectclock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please send your openion to me
'My address is dundee_united70@hotmail.com
'I believe that u will like this software
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GCL_HCURSOR = (-12)
Private hOldCursor As Long
Dim X As Integer
Dim Y As Integer
Dim z As Integer
Dim i As Integer
Dim PI, Radius, Radians As Double
Dim hrs, Min, Sec  As Integer
Dim Pq As Double
Dim b As Integer
Dim str As String
Sub cur()
Dim hNewCursor As Long
hNewCursor = LoadCursorFromFile(App.Path & "\n.ani")
hOldCursor = SetClassLong(hwnd, GCL_HCURSOR, hNewCursor)
End Sub

Private Sub b_Click(Index As Integer)
Image1.Picture = Picture3.Picture
End Sub

Private Sub bl_Click(Index As Integer)
Image1.Picture = Picture3.Picture
End Sub

Private Sub Form_Load()
PI = 3.14159265358979
seconds.X1 = 2660
seconds.Y1 = 3165
minutes.X1 = seconds.X1
hours.X1 = seconds.X1
minutes.Y1 = seconds.Y1
hours.Y1 = seconds.Y1
Radius = 1680
str = "Perfect clock by dundee_united70@hotmail.com"
b = Len(str)
i = 1
Call cur
End Sub

Private Sub gr_Click(Index As Integer)
Image1.Picture = Picture4.Picture
End Sub

Private Sub r_Click(Index As Integer)
Image1.Picture = Picture1.Picture
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 1000
hrs = Hour(Time$) * 5 - 12
Sec = Second(Time$)
Min = Minute(Time) - 14
sndPlaySound (App.Path & "\t.wav"), 1
Radians = Sec / 180 * PI
seconds.X2 = seconds.X1 + Cos(Radians * 6) * (Radius - 600)
seconds.Y2 = seconds.Y1 + Sin(Radians * 6) * (Radius - 600)
Radians = Min / 180 * PI
minutes.X2 = seconds.X1 + Cos(Radians * 6) * (Radius - 900)
minutes.Y2 = seconds.Y1 + Sin(Radians * 6) * (Radius - 900)
Radians = hrs / 180 * PI
hours.X2 = seconds.X1 + Cos(Radians * 6) * (Radius - 1200)
hours.Y2 = seconds.Y1 + Sin(Radians * 6) * (Radius - 1200)
End Sub

Private Sub Timer2_Timer()
Timer2.Interval = 100
Me.Caption = Left(str, i)
i = i + 1
If i = b + 1 Then
Timer2.Interval = 3000
i = 0
End If
End Sub

Private Sub v_Click(Index As Integer)
Image1.Picture = Picture2.Picture
End Sub

Private Sub ye_Click(Index As Integer)
Image1.Picture = Picture5.Picture
End Sub
