VERSION 5.00
Begin VB.UserControl UserControl1 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2820
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   ForwardFocus    =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   2820
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2850
      Top             =   2100
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2610
      Top             =   1050
   End
   Begin VB.Image Image2 
      Height          =   345
      Index           =   0
      Left            =   1350
      Picture         =   "UserControl21.ctx":0000
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1050
   End
   Begin VB.Image Image3 
      Height          =   345
      Index           =   0
      Left            =   300
      Picture         =   "UserControl21.ctx":24A32
      Stretch         =   -1  'True
      Top             =   5130
      Width           =   930
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   510
      TabIndex        =   2
      Top             =   1950
      Width           =   1905
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   660
      TabIndex        =   0
      Top             =   945
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   660
      TabIndex        =   1
      Top             =   930
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   660
      Picture         =   "UserControl21.ctx":49464
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1140
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ItsClicked As Boolean
Public TextLabel As String
Public Identifier As Integer
Public STyleIndex As Integer
'I didn't add code to explain this control because
'I though that people aren't interested in how the fancy
'buttons work!! I will post some code up at
'"planet source code" much later in the year (when my
'exams are over).
'
'If you want to know more, I can tell you
'That I got the ideas from this,
'
'Search for a .Manifest file in your windows or system32
'directory and make a copy of it.
'Get a vb project, and copy the whole name of the file to
'the clipboard ie "My Program.exe" and move the .manifest
'to the same location as the "my program.exe".
'now rename the .manifest file to what ever the project's
'file name is ie "my program.exe.manifest"
'run My program.exe and observe the buttons and textboxes
'Have Fun!!
Private Sub Image1_DblClick()
Image1_Click
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Picture = Image2(STyleIndex).Picture Then Exit Sub
Image1.Picture = Image2(STyleIndex).Picture
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Picture = Image3(STyleIndex).Picture Then Exit Sub
Image1.Picture = Image3(STyleIndex).Picture
End Sub

Private Sub Label1_Click()
Image1_Click
End Sub

Private Sub Label1_DblClick()
Image1_Click
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y <= 30 Then Y = 31
If Y >= Label1.Height - 30 Then Y = Image1.Height
Call image1_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label2_DblClick()
Image1_Click
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y <= 30 Then Y = 31
If Y >= Label1.Height - 30 Then Y = Image1.Height
Call image1_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y >= Image1.Height - 30 Or Y <= 30 Or X <= 30 Or X >= Image1.Width - 30 Then
    Timer2_Timer
    Exit Sub
End If

If Label1.ForeColor = 14737632 Then Exit Sub
Label2.ForeColor = 14737632
Label1.ForeColor = 4210752
Timer2.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Label2_Click()
Image1_Click
End Sub

Public Sub Image1_Click()
Form1.haSBeenclicked (Identifier)
End Sub

Private Sub label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = Image2(STyleIndex).Picture
End Sub
Private Sub label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = Image3(STyleIndex).Picture
End Sub
Private Sub label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = Image2(STyleIndex).Picture
End Sub
Private Sub label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = Image3(STyleIndex).Picture
End Sub

Public Sub Label3_Click()
Image1_Click
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer2_Timer
End Sub

Public Sub Timer1_Timer()
ItsClicked = False
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Label2.ForeColor = 12632256
Label1.ForeColor = 4210752
Timer2.Enabled = False
End Sub

Private Sub UserControl_Initialize()
usercontrol_Resize
End Sub

Private Sub usercontrol_Resize()
Image1.Top = 0
Image1.Left = 0
Image1.Width = UserControl.Width
Image1.Height = UserControl.Height
Label1.Top = (UserControl.Height / 2) - 100 - 15
Label1.Width = UserControl.Width
Label1.Left = 0
Label2.Top = (UserControl.Height / 2) - 100
Label2.Width = UserControl.Width
Label2.Left = 0
Label1.Height = UserControl.Height - Image1.Top
Label2.Height = UserControl.Height - Image1.Top
Label1.Caption = TextLabel
Label2.Caption = TextLabel
Label3.Top = UserControl.Height - 30
Label3.Left = 0
Label3.Width = UserControl.Width
'Image2.Top = UserControl.Height + 100
'Image3.Top = UserControl.Height + 100
Image1.Picture = Image3(STyleIndex).Picture
End Sub

Public Sub ChangeLabel(NewLabel As String)
Label1.Caption = NewLabel
Label2.Caption = NewLabel
End Sub


Public Sub ChangesTYLE(STyleIndex2 As Integer)
STyleIndex = STyleIndex2
Image1.Picture = Image3(STyleIndex).Picture
End Sub

