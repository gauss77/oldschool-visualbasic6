VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set & Destroy®"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl MMC 
      Height          =   375
      Left            =   4680
      TabIndex        =   20
      Top             =   8160
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      PlayEnabled     =   -1  'True
      DeviceType      =   ".wav"
      FileName        =   "D:\VB6\Programming I\Set & Destroy\EXPLODE.WAV"
   End
   Begin VB.CheckBox chkGrid 
      Caption         =   "Set Grid Scale"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      ToolTipText     =   "Enables grid scale"
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      DragIcon        =   "frmGame.frx":1272
      Height          =   615
      Left            =   7200
      Picture         =   "frmGame.frx":16B4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Help!"
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton cmdAgain 
      Caption         =   "Play Again"
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Click to play again"
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox txtSpeed 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "100"
      ToolTipText     =   "Enter speed of shot"
      Top             =   7920
      Width           =   975
   End
   Begin VB.CheckBox chkStartorStop 
      Caption         =   "Start or Stop"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Launch the bomb"
      Top             =   8280
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6000
      Top             =   7200
   End
   Begin VB.CommandButton cmdDone 
      Cancel          =   -1  'True
      Caption         =   "Done"
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Click to exit"
      Top             =   8040
      Width           =   1575
   End
   Begin VB.TextBox txtVertical 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "45"
      ToolTipText     =   "Enter angle in degrees"
      Top             =   7320
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   6975
      Left            =   240
      ScaleHeight     =   6915
      ScaleWidth      =   9945
      TabIndex        =   0
      Top             =   0
      Width           =   10000
      Begin VB.Shape shpTarget 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   7680
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   735
      End
      Begin VB.Line LineX 
         Index           =   6
         X1              =   0
         X2              =   9960
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Shape shpBomb 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         Shape           =   3  'Circle
         Top             =   6600
         Width           =   255
      End
      Begin VB.Line LineY 
         Index           =   41
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineX 
         Index           =   29
         X1              =   0
         X2              =   9960
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Line LineX 
         Index           =   28
         X1              =   0
         X2              =   9960
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line LineX 
         Index           =   27
         X1              =   0
         X2              =   9960
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line LineX 
         Index           =   26
         X1              =   0
         X2              =   9960
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line LineX 
         Index           =   25
         X1              =   0
         X2              =   9960
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line LineX 
         Index           =   24
         X1              =   0
         X2              =   9960
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line LineX 
         Index           =   23
         X1              =   0
         X2              =   9960
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line LineX 
         Index           =   22
         X1              =   0
         X2              =   9960
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line LineX 
         Index           =   21
         X1              =   0
         X2              =   9960
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line LineX 
         Index           =   20
         X1              =   0
         X2              =   9960
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line LineX 
         Index           =   19
         X1              =   0
         X2              =   9960
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line LineX 
         Index           =   18
         X1              =   0
         X2              =   9960
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line LineX 
         Index           =   17
         X1              =   0
         X2              =   9960
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line LineX 
         Index           =   16
         X1              =   0
         X2              =   9960
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line LineX 
         Index           =   15
         X1              =   0
         X2              =   9960
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line LineX 
         Index           =   14
         X1              =   0
         X2              =   9960
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line LineX 
         Index           =   13
         X1              =   0
         X2              =   9960
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line LineX 
         Index           =   12
         X1              =   0
         X2              =   9960
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line LineX 
         Index           =   11
         X1              =   0
         X2              =   9960
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line LineX 
         Index           =   10
         X1              =   0
         X2              =   9960
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line LineX 
         Index           =   9
         X1              =   0
         X2              =   9960
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line LineX 
         Index           =   8
         X1              =   0
         X2              =   9960
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line LineX 
         Index           =   7
         X1              =   0
         X2              =   9960
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line LineX 
         Index           =   5
         X1              =   0
         X2              =   9960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line LineX 
         Index           =   4
         X1              =   0
         X2              =   9960
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line LineX 
         Index           =   3
         X1              =   0
         X2              =   9960
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line LineX 
         Index           =   2
         X1              =   0
         X2              =   9960
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line LineX 
         Index           =   1
         X1              =   0
         X2              =   9960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line LineY 
         Index           =   40
         X1              =   9600
         X2              =   9600
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   39
         X1              =   9360
         X2              =   9360
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   38
         X1              =   9840
         X2              =   9840
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   37
         X1              =   8880
         X2              =   8880
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   36
         X1              =   9120
         X2              =   9120
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   35
         X1              =   7200
         X2              =   7200
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   34
         X1              =   7440
         X2              =   7440
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   33
         X1              =   7680
         X2              =   7680
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   32
         X1              =   7920
         X2              =   7920
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   31
         X1              =   8160
         X2              =   8160
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   30
         X1              =   8400
         X2              =   8400
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   29
         X1              =   6960
         X2              =   6960
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   28
         X1              =   6720
         X2              =   6720
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   27
         X1              =   6480
         X2              =   6480
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   26
         X1              =   6240
         X2              =   6240
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   25
         X1              =   6000
         X2              =   6000
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   24
         X1              =   5760
         X2              =   5760
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   23
         X1              =   5520
         X2              =   5520
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   22
         X1              =   5280
         X2              =   5280
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   21
         X1              =   8640
         X2              =   8640
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   20
         X1              =   4800
         X2              =   4800
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   19
         X1              =   4560
         X2              =   4560
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   18
         X1              =   4320
         X2              =   4320
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   17
         X1              =   5040
         X2              =   5040
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   16
         X1              =   2880
         X2              =   2880
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   15
         X1              =   3120
         X2              =   3120
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   14
         X1              =   3360
         X2              =   3360
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   13
         X1              =   3600
         X2              =   3600
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   12
         X1              =   3840
         X2              =   3840
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   11
         X1              =   2640
         X2              =   2640
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   10
         X1              =   2400
         X2              =   2400
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   9
         X1              =   2160
         X2              =   2160
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   8
         X1              =   1920
         X2              =   1920
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   7
         X1              =   1680
         X2              =   1680
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   6
         X1              =   1440
         X2              =   1440
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   5
         X1              =   1200
         X2              =   1200
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   4
         X1              =   960
         X2              =   960
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   3
         X1              =   720
         X2              =   720
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   2
         X1              =   480
         X2              =   480
         Y1              =   -120
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   1
         X1              =   240
         X2              =   240
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineY 
         Index           =   0
         X1              =   4080
         X2              =   4080
         Y1              =   0
         Y2              =   6840
      End
      Begin VB.Line LineX 
         Index           =   0
         X1              =   0
         X2              =   9960
         Y1              =   6840
         Y2              =   6840
      End
   End
   Begin VB.Label lblNumTries 
      Height          =   255
      Left            =   4080
      TabIndex        =   19
      Top             =   8040
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Number of tries:"
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   17
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "Angle in radians:"
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Speed:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Y:"
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "X:"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label lblRadians 
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblY 
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label lblX 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Enter angle:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   7080
      Width           =   975
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sami Eljabali
'Programming I
'Semester Project
Option Explicit
Private Numtries As Integer

'  SSS    EEEEEE   TTTTTTTT    &      DDDDD      EEEE    SSS  TTTTTT   RRRRR       OOOOOO     Y     Y
' S       E           T        &      D    D     E      S        T     R    R     O      O    Y     Y
'  S      E           T        &      D     D    E      S        T     R    R    O        O   Y     Y
'   S     E           T       &&&&&   D      D   E       S       T     R    R    O        O    Y   Y
'    S    E           T      &        D      D   E        S      T     RRRRRR    O        O     YYY
'     S   EEEEE       T       &&&&&   D      D   EEEE       S    T     R     R   O        O      Y
'     S   E           T      &        D      D   E          S    T     R     R   O        O      Y
'     S   E           T       &&&&&   D      D   E          S    T     R     R   O        O      Y
'    S    E           T        &      D      D   E         S     T     R     R    O       O      Y
'   S     E           T        &      D     D    E        S      T     R     R     O     O       Y
'SSS      EEEEEEE     T        &      DDDDDD     EEEE  SS        T     R     R      OOOOO        Y

'                                 SET & DESTROY®
'**********************************************************************************
'* Welcome to the coding of the game of Set & Destroy®. I have programmed the game*
'* to calculate the user's input of the angle in degrees then converted into      *
'* radians so that Visual Basic can calculate the cordinates of the moving ball.  *
'* I have also set boarders to the game to get rid of bugs and etc. There is also *
'* a grid scale system to make the aiming of the bomb easier.The number of tries  *
'* is counted from the number of tries since you won till you win again you may   *
'* enter the same coordinates over and over to till you get it to be fair.        *
'* Preconditions: Enter angle and speed                                           *
'* Postconditions: Enter speed from 1-500. Enter angle from -360 to 360           *
'* (Numbers greater or lower can cause overflow!)                                 *
'**********************************************************************************
Private Sub Timer1_Timer()


    'The game starts when check box is unchecked
    If chkStartorStop.Value = vbChecked Then Exit Sub
    
    Dim intSound As Integer
    
    chkStartorStop.Value = vbGrayed 'Disables cheating(To chage entered data)
    Dim X As Long
    Dim Y As Long
    Dim Angle As Double
    Dim Speed As Long
    Dim YorN As String

    
    'Converting degrees to radians
    Angle = txtVertical / 180
    Angle = Angle * 3.14159
    lblRadians.Caption = Angle 'Displays angle in radians
    
    Speed = CLng(txtSpeed)
    
    'Getting the X cordinates
    X = Speed * Cos(Angle)
    
    'Getting the Y cordinates
    Y = Speed * Sin(Angle)
    
    'Making the ball move smoother *Got help
    shpBomb.Visible = False
    shpBomb.Top = shpBomb.Top - Y
    'Displaying Y cordinates
    lblY.Caption = shpBomb.Top
    shpBomb.Left = shpBomb.Left + X
    'Displaying X cordinates
    lblX.Caption = shpBomb.Left
    shpBomb.Visible = True
    
    'Displaying the number of tries at all times
    lblNumTries = Numtries
    
    'Setting boarders
    Call Boarders(YorN)
        
    'Enabling hitting the target
    Call Target(Numtries)
End Sub

Sub Boarders(ByVal YorN As String)
'Setting boarders of the width
    If shpBomb.Left > 9720 Or shpBomb.Left < 0 Then
        chkStartorStop = vbChecked 'Stops movement
        'Enable sound of missing
        MMC.DeviceType = "WaveAudio"
        MMC.FileName = "Lost.wav"
        MMC.Command = "Open"
        MMC.Command = "Play"
        MMC.PlayEnabled = True

        MsgBox "You missed."
        Do Until YorN = "y" Or YorN = "n"
            YorN = InputBox("Do you want to play again? y or n?", "Set & Destroy®")
            If YorN = "y" Then
                Unload Me 'Leaves one form insight
                frmIntro.Show 'Back to main menu
            ElseIf YorN = "n" Then
                Unload Me
                frmEnding.Show 1 'Makes concentration on the form only*Got help
                Exit Sub
            End If
        Loop
    End If

'Setting boarders of the height
    If shpBomb.Top > 6600 Or shpBomb.Top < 0 Then
        chkStartorStop = vbChecked 'Stops movement
        'Enable sound of missing
        MMC.DeviceType = "WaveAudio"
        MMC.FileName = "Lost.wav"
        MMC.Command = "Open"
        MMC.Command = "Play"
        MMC.PlayEnabled = True

        MsgBox "You missed."

        Do Until YorN = "y" Or YorN = "n"
            YorN = InputBox("Do you want to play again? y or n?", "Set & Destroy®")
            If YorN = "y" Then
                frmIntro.Show 'Back to main menu
                Unload Me 'Leaves one form insight
            ElseIf YorN = "n" Then
                Unload Me
                frmEnding.Show 1 'Makes concentration on the form only*Got help
                Exit Sub
            End If
        Loop
    End If
End Sub

Sub Target(ByRef Numtries As Integer)
     'If the bomb hits the target
    Dim MsgYorN As String
    If (shpBomb.Left >= shpTarget.Left + 13 And shpBomb.Left <= (shpTarget.Left + shpTarget.Width) Or shpBomb.Left <= shpTarget.Width) And _
    (shpBomb.Top >= shpTarget.Top And shpBomb.Top <= shpTarget.Top + shpTarget.Height) Then
      'Then...
        
        'Enabling sound of hitting the target
        MMC.DeviceType = "WaveAudio"
        MMC.FileName = "EXPLODE.wav"
        MMC.Command = "Open"
        MMC.Command = "Play"
        MMC.PlayEnabled = True

        chkStartorStop = vbChecked 'Stop movement
        shpBomb.Visible = False 'The bomb can no longer be seen
        shpTarget.BackColor = vbRed 'The target turns red
            If Numtries = 1 Then 'If got it at first time then
                MsgBox "WOW! Got it on your 1st try!!" 'this msgbox appears
            Else
                MsgBox "It took you" & " " & Numtries & " " & "tries."
            End If
        'Msgbox asking if you want to play again or not
        MsgYorN = MsgBox("You got it. Want to play again?", vbYesNo, "Set & Destroy®") 'Message box appears
        If MsgYorN = vbYes Then
           frmIntro.Show
           Numtries = 0 'Makes the number of tries 0 when user wins
           Unload Me
        ElseIf MsgYorN = vbNo Then
           frmEnding.Show
           Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    MMC.Visible = False 'Disabling the view of media player
    MMC.FileName = ""
    
    Dim intX As Long
    Dim intY As Long
    
    'Entering the coordinates of the target
    'X coordinate
    intX = InputBox("Enter the X cordinnate of your target from (0-9120)", "Set & Destroy®", "1234")
    Do Until intX >= 0 And intX <= 9120
        intX = InputBox("Enter the X cordinnate of your target from (0-9120)", "Set & Destroy®", "1234")
    Loop
    shpTarget.Left = intX
    'Y coordinate
    intY = InputBox("Enter the Y cordinnate of your target from (0-6240)", "Set & Destroy®", "1234")
    Do Until intY >= 0 And intY <= 6240
        intY = InputBox("Enter the Y cordinnate of your target from (0-6240)", "Set & Destroy®", "1234")
    Loop
    shpTarget.Top = intY
    
    'Getting number of tries
    Numtries = Numtries + 1
    
    'Disabling the grid scale at form load
    Dim intYNum As Integer
    Dim intXNum As Integer
    
    If chkGrid.Value = False Then
        For intYNum = 0 To 41
         LineY(intYNum).Visible = False
        Next intYNum
            For intXNum = 0 To 29
              LineX(intXNum).Visible = False
            Next intXNum
    End If
End Sub

Private Sub chkGrid_Click()
    'Setting the grid sclae
    Dim intYNum As Integer
    Dim intXNum As Integer
    
    'Enable
    If chkGrid = vbChecked Then
        For intYNum = 0 To 41
            LineY(intYNum).Visible = True
        Next intYNum
            For intXNum = 0 To 29
                LineX(intXNum).Visible = True
            Next intXNum
    Else 'Disable
        For intYNum = 0 To 41
            LineY(intYNum).Visible = False
        Next intYNum
            For intXNum = 0 To 29
                LineX(intXNum).Visible = False
            Next intXNum
    End If
End Sub

Private Sub cmdAgain_Click()
    'Takes you to the main menu
    frmIntro.Show
    'This is done to leave only one form insight
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    'This pauses the game when entering the help wizard
    chkStartorStop.Value = vbChecked
    
    'Displays the help wizard
    'The "1" after ".show" makes the concentration only on the form itself*Got help
    frmHelp.Show 1
End Sub

Private Sub cmdAgain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If over the button(Dark Grey)
    cmdAgain.BackColor = &HC0C0C0
End Sub

Private Sub cmdDone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If over the button(Dark Grey)
    cmdDone.BackColor = &HC0C0C0
End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If over the button(Dark Grey)
    cmdHelp.BackColor = &HC0C0C0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If not over button(s)(Grey)
    If cmdAgain.BackColor = &HC0C0C0 Then
        cmdAgain.BackColor = &H8000000F
    End If
    
    If cmdDone.BackColor = &HC0C0C0 Then
        cmdDone.BackColor = &H8000000F
    End If
    
    If cmdHelp.BackColor = &HC0C0C0 Then
        cmdHelp.BackColor = &H8000000F
    End If
End Sub

Private Sub cmdDone_Click()
    'This pauses the game when leaving the game
    chkStartorStop.Value = vbChecked
    'Takes you to credits
    'The "1" after ".show" makes the concentration only on the form itself*Got help
     frmEnding.Show 1
    'This is done to leave only one form insight
    Unload Me
End Sub
