VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Bassmaker"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4020
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   30
      ScaleHeight     =   615
      ScaleWidth      =   3975
      TabIndex        =   12
      Top             =   3480
      Width           =   3975
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0 Hertz"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   2790
         TabIndex        =   13
         Top             =   420
         Width           =   1155
      End
   End
   Begin Bass_Maker.UserControl1 UserControl15 
      Height          =   345
      Left            =   3450
      TabIndex        =   11
      Top             =   1050
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   609
   End
   Begin Bass_Maker.UserControl1 UserControl14 
      Height          =   345
      Left            =   1710
      TabIndex        =   10
      Top             =   1050
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   609
   End
   Begin Bass_Maker.UserControl1 UserControl11 
      Height          =   345
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
   End
   Begin VB.CommandButton Command5 
      Caption         =   "New"
      Height          =   315
      Left            =   3450
      TabIndex        =   6
      ToolTipText     =   "Clear current project"
      Top             =   1080
      Width           =   525
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load Project"
      Height          =   315
      Left            =   1710
      TabIndex        =   5
      ToolTipText     =   "Load a previously saved project"
      Top             =   1080
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form1.frx":08CA
      ToolTipText     =   "Frequency,length [enter] frequency,length [Enter] and so on. Frequency is not in Hertz but length is in amount of wavelengths"
      Top             =   1440
      Width           =   3885
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3540
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Wave"
      Height          =   345
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Make a wave file"
      Top             =   60
      Width           =   1995
   End
   Begin Bass_Maker.UserControl1 UserControl12 
      Height          =   345
      Left            =   2070
      TabIndex        =   8
      Top             =   60
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   609
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Leave"
      Height          =   345
      Left            =   2070
      TabIndex        =   1
      ToolTipText     =   "Exit the program"
      Top             =   60
      Width           =   1845
   End
   Begin Bass_Maker.UserControl1 UserControl13 
      Height          =   345
      Left            =   30
      TabIndex        =   9
      Top             =   1050
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save Project"
      Height          =   315
      Left            =   30
      TabIndex        =   4
      ToolTipText     =   "Save your project"
      Top             =   1080
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":51D2
      Height          =   615
      Left            =   30
      TabIndex        =   2
      ToolTipText     =   "This tells you how you should set out your projects"
      Top             =   450
      Width           =   3915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Y1#
Dim Place$
Dim tempF50$
Dim temp$
Dim Y2$
Dim Y3$
Dim Y4$
Dim T#
Dim FileSizeE&

Private Sub Command1_Click()
'If an error occurs then exit
On Error Resume Next
'Close any open files
Close
'call the mess subroutine to display "loading"
Mess ("Gathering Information and Loading Files...")
'open the commondialog to select where to save the file
CommonDialog1.FileName = "SaveToWave.wav"
CommonDialog1.ShowSave
'open the wave file to be saved
Open CommonDialog1.FileName + ".wav" For Output As #1
'Save the contents of the textbox to a file, makes it
'easier to manipulate and load but is slower than working
'with a string (I'm just lazy)
Open "C:\temp.dat" For Output As #4
Print #4, Text1.Text
Close #4
'Inform the user that is is calculating the header
Mess ("Calculating Header Information...")
'open the temp file we just saved (i.e. text1.text)
Open "C:\temp.dat" For Input As #3
'Print the first 4 bytes of the header (all waves start
'with RIFF, even some Avi file do too!)
'-------------
'the filesizee& is used to calculate the length of the
'wave file in the header (is needed)
Print #1, "RIFF";: FileSizeE& = 110
'read each frequency and amount of waves from temp.dat
While Not EOF(3)
    'get each frequency and amount of waves
    Input #3, X5$: Input #3, X6$: T# = Val(X5$)
    'if frquency is 0, then it assumes it is the end of
    'the file
    If T# = 0 Then GoTo 204
    'used to calculate exacally how big the file will be,
    'NOTE:- this part does not create any part of the
    'wave file, just calculating its size!!
    X4# = Val(X6$) * 2
        For O# = 0 To (3.14159 * X4#) Step T#
            FileSizeE& = FileSizeE& + 4
        Next O#
Wend
204
'Saves the expected size of the wave file into the header
'by converting the file length to hex (tempf50$),
'splitting the hex number up into groups of 2 digits,
'converting the hex number back to decimal and convert the
'value into a character (i.e. Chr$())
Mess ("Saving Header Information...")
tempF50$ = Hex(FileSizeE&): tempF50$ = Space(8 - Len(tempF50$)) + tempF50$: HexConverTT (Mid$(tempF50$, 7, 2)): place1$ = Place$: HexConverTT (Mid$(tempF50$, 5, 2)): place2$ = Place$: HexConverTT (Mid$(tempF50$, 3, 2)): place3$ = Place$: HexConverTT (Mid$(tempF50$, 1, 2)): Print #1, place1$ + place2$ + place3$ + Place$;
Print #1, Chr$(&H57) + Chr$(&H41) + Chr$(&H56) + Chr$(&H45) + Chr$(&H66) + Chr$(&H6D) + Chr$(&H74) + Chr$(&H20) + Chr$(&H10) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) + Chr$(&H1) + Chr$(&H0) + Chr$(&H2) + Chr$(&H0) + Chr$(&H22) + Chr$(&H56) + Chr$(&H0) + Chr$(&H0) + Chr$(&H88) + Chr$(&H58) + Chr$(&H1) + Chr$(&H0) + Chr$(&H4) + Chr$(&H0) + Chr$(&H10) + Chr$(&H0) + Chr$(&H64) + Chr$(&H61) + Chr$(&H74) + Chr$(&H61);
tempF50$ = Hex(FileSizeE& - 110): tempF50$ = Space(8 - Len(tempF50$)) + tempF50$: HexConverTT (Mid$(tempF50$, 7, 2)): place1$ = Place$: HexConverTT (Mid$(tempF50$, 5, 2)): place2$ = Place$: HexConverTT (Mid$(tempF50$, 3, 2)): place3$ = Place$: HexConverTT (Mid$(tempF50$, 1, 2)): Print #1, place1$ + place2$ + place3$ + Place$;
'Display's message creating wave file
Mess ("Creating Wave File... Estimated" + Str$(Int((FileSizeE& / 5292000) * 100) * 0.6) + " Second File... " + Str$(Int(FileSizeE& / 10000) / 100) + "Mb To Be Written... Please Wait...")
'close temp.dat and open it again from the start,
'otherwise it will try and start reading from the end
'of temp.dat (see above in header calculation)
Close #3: Open "C:\temp.dat" For Input As #3
'This is the loop where it all happens, it generates raw
'data in wave format (i.e. a normal wave file without a
'header)
While Not EOF(3)
    'read each frequency and amount of waves from temp.dat
    Input #3, X5$: Input #3, X6$: T# = Val(X5$)
    'Assumes that 0 freq is EOF (end of file)
    If T# = 0 Then GoTo 203
    'using the sin function is generates a sine wave!!
    X4# = Val(X6$) * 2
        'steps a full 2*pi radians X4# times
        For O# = 0 To (3.14159 * X4#) Step T#
            'A number that when multiplied with the
            'radians, seemed to work. otherwise it
            'spat out static!!
            Y1# = 30000 * Sin(O#)
            If Int(Y1#) <= 0 Then Y1# = 65535 + Y1#
            'Converts the number to Hex then splits it
            'to groups of 2 (i.e LLLLRRRR where L=left
            'and R=right channel, 16-bit wave format)
            'then converts back to decimal and then saves
            'as a character
            Y2$ = Hex(Int(Y1#)): Y2$ = Space(4 - Len(Y2$)) + Y2$: Y3$ = Mid$(Y2$, 1, 2): Y4$ = Mid$(Y2$, 3, 2): HexConverTT (Y4$)
            'Saves Raw Wave data to file
            Print #1, Place$;: temp$ = Place$: HexConverTT (Y3$): Print #1, Place$ + temp$ + Place$;
        Next O#
Wend
'Message nearly finished
203 Mess ("Performing Final Steps...")
'Creates a little comment at end of wave file (sort of
'like a ID3 tag in Mp3's (if you do change it, make sure
'it is the same charater lengh!!
message1$ = "Made With Bass Maker, Alex Murray": Print #1, Chr$(&H4C) + Chr$(&H49) + Chr$(&H53) + Chr$(&H54) + Chr$(&H42) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) + Chr$(&H49) + Chr$(&H4E) + Chr$(&H46) + Chr$(&H4F) + Chr$(&H49) + Chr$(&H53) + Chr$(&H46) + Chr$(&H54) + Chr$(&H35) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) + message1$ + Space(52 - Len(message1$)) + Chr$(&H0) + Chr$(&H0);
'Close all open files
Close
'Done!!!!
10 Call MsgBox("Done", vbInformation, "Done")
Label1.Caption = "Enter a number between say 0.01 and 0.03 for bass and a integer for the amount of wave lengths, E.g. 0.025, 2, you can repeat this as shown below"
End Sub

Private Sub HexConverTT(Hexd$)
'convert hex to decimal!! (I don't know the command to
'convert hex to decimal so....)
A1$ = Mid$(Hexd$, 1, 1): B1$ = Mid$(Hexd$, 2, 1)
'if digit is a letter, then convert it to a number
'if Hex=base 16, i.e. 0123456789ABCDEF
If UCase$(A1$) = " " Then A1$ = "0"
If UCase$(B1$) = " " Then B1$ = "0"
If UCase$(A1$) = "A" Then A1$ = "10"
If UCase$(A1$) = "B" Then A1$ = "11"
If UCase$(A1$) = "C" Then A1$ = "12"
If UCase$(A1$) = "D" Then A1$ = "13"
If UCase$(A1$) = "E" Then A1$ = "14"
If UCase$(A1$) = "F" Then A1$ = "15"
If UCase$(B1$) = "A" Then B1$ = "10"
If UCase$(B1$) = "B" Then B1$ = "11"
If UCase$(B1$) = "C" Then B1$ = "12"
If UCase$(B1$) = "D" Then B1$ = "13"
If UCase$(B1$) = "E" Then B1$ = "14"
If UCase$(B1$) = "F" Then B1$ = "15"
'if last digit is zero then skip this step (or say 01 will
'become 16 instead of 1 in base 10
If B1$ = " " Then GoTo 10
'multiplys by the power corresponding to its position
'i.e a number in position 3 from left is raised to power
'32, a number in position 2 from left is raised to power
'16, a number in position 4 from left is raised to power
'48 and so on and then all digits are added i.e. position
'1 + position 2 and so on (find hex conversions on the
'net)
d1% = (Val(A1$) * 16) + Val(B1$): GoTo 20
10 d1% = Val(A1$)
'return the decimal value
20 Place$ = Chr$(d1%)
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
On Error Resume Next
'saves the contents of text1.text to a .txt file so you
'can save your projects
CommonDialog1.FileName = "SaveToText.txt"
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then Exit Sub
Open CommonDialog1.FileName + ".Txt" For Output As #4
Print #4, Text1.Text
Close
End Sub

Private Sub Command4_Click()
On Error Resume Next
'loads the contents of a .txt file to text1.text so you
'can load your projects
CommonDialog1.FileName = "LoadFromText.txt"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
'If you havn't seen anyone use Binary to open files yet,
'It is the fastest method of opening files without any
'complex code!! First type a length in bytes to be read
'into a string (spaces will do), then use the get function
'as you would in random, ie get #2,2,Temp$ is get bytes
'1*len(temp$) to 2*len(temp$). here I open the whole file
'with one read from the file
template1$ = Space(FileLen(CommonDialog1.FileName))
Open CommonDialog1.FileName For Binary As #5
Get #5, 1, template1$
Close
'saves data to text1.text
Text1.Text = template1$
End Sub

Private Sub Command5_Click()
Text1.Text = ""
End Sub


Private Sub Form_Load()
'These are my nifty user controls
'if you don't under stand them fully then don't worry!!
'My code for these is a mess because I just used a lot
'of guessing and luck to make them work!!
UserControl11.ChangeLabel ("Make Wave")
'UserControl11.ChangesTYLE (4)
UserControl11.Identifier = 1
UserControl12.ChangeLabel ("Leave Bass Maker")
'UserControl12.ChangesTYLE (4)
UserControl12.Identifier = 2
UserControl13.ChangeLabel ("Save Project")
'UserControl13.ChangesTYLE (4)
UserControl13.Identifier = 3
UserControl14.ChangeLabel ("Load Project")
'UserControl14.ChangesTYLE (4)
UserControl14.Identifier = 4
UserControl15.ChangeLabel ("New")
'UserControl15.ChangesTYLE (4)
UserControl15.Identifier = 5
End Sub

Public Sub haSBeenclicked(Identifier As Integer)
'if you click one of my usercontrols, then run the requred
'subroutine
If Identifier = 1 Then Command1_Click
If Identifier = 2 Then Command2_Click
If Identifier = 3 Then Command3_Click
If Identifier = 4 Then Command4_Click
If Identifier = 5 Then Command5_Click
End Sub

Private Sub Form_Resize()
'on form resize, re-arrange and resize the contents in the
'form
On Error Resume Next
Me.Width = 4140
Text1.Height = Me.Height - (4605 - 1995)
Picture1.Top = Me.Height - (4605 - 3480)
End Sub


Private Sub Mess(Mess1$)
'Display a message (reduces code)
Label1.Caption = Mess1$
Label1.Refresh
Me.Refresh
End Sub

Private Sub Timer1_Timer()
'This is the little splash screen
Load Form2
Form2.Show
Form2.SetFocus
Timer1.Enabled = False
End Sub

Private Sub Text1_Click()
On Error Resume Next
'This form is the little subroutine that draws the
'sine wave in the black picture box
A$ = Text1.SelText
'clear the picture box
Picture1.Cls
If Val(A$) = 0 Then Exit Sub
'i don't know why this is here twice, but I'm sure there
'was a reason for it??
Picture1.Cls
'get a resonable value to work with as a integer (note the
'use of longint i.e. '&', mainly because I forgot about '#'
'at the time)
temp55& = Val(A$) * 10000 / 2.5
'For certain cases, error occur, These are just displayed
'as hertz too high or too low
If temp55& < 12000 Then
    If temp55& <= 0 Then
        Label2.Caption = "Hertz Too Low"
        Exit Sub
    Else
        Label2.Caption = Str$(temp55&) + " Hertz"
    End If
Else
    Label2.Caption = "Hertz Too High"
    Exit Sub
End If
'Draw the line half way between the top and bottom of the
'picture box
Picture1.Line (0, Picture1.Height / 2 - 30)-(Picture1.Width, Picture1.Height / 2 - 30)
109
        'Draw the sine wave, go from 0 to 2*pi steping
        'value of the freq, ie lower freq, bigger waves
        For O# = 0 To (3.14159 * 2) Step Val(A$)
            Y1# = (Picture1.Height - 70) * Sin(O#) / 2
            Y23# = (Picture1.Height - 70) * Sin(O# + Val(A$)) / 2
            temp33& = temp33& + 2
            'draw the wave
            Picture1.Line (temp33&, (Picture1.Height / 2) + Y1#)-(temp33& + 15, (Picture1.Height / 2 + Y23#))
        Next O#
        'if the wave is sufficiently big, then stop
        'drawing waves
        If temp33& < Picture1.Width Then GoTo 109
End Sub

Private Sub Text1_DblClick()
'double click it same as single click
Text1_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'get the value of the selected text and draw a wave
'(only if you select something)
A$ = Text1.SelText
If Val(A$) > 0 Then Text1_Click
End Sub
