VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   1500
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   930
      Top             =   30
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   30
      Top             =   30
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   480
      Top             =   30
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2610
      TabIndex        =   0
      Top             =   1230
      Width           =   1485
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   2610
      X2              =   4440
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Don't worry too much about this form,
'It is just for looks :) and can be deleted safely!!
'(remember to change the startup object in project
'then properties (at bottom of menu)
Private Sub Form_Load()
Line1.X2 = Line1.X1
End Sub

Private Sub Timer1_Timer()
Me.Hide
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Load Form1
Form1.Show
Me.Show
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
'the nifty counter bar thingy
Line1.X2 = Line1.X2 + 100
If Line1.X2 >= 4340 Then
    Line1.Visible = False
    Label1.Visible = False
    Timer3.Enabled = False
End If
End Sub
