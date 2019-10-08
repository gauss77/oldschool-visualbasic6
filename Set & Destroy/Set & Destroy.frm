VERSION 5.00
Begin VB.Form frmIntro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set & Destroy®"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12120
   Icon            =   "Set & Destroy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   12120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Cancel          =   -1  'True
      Caption         =   "Done"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click to exit"
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Begin"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click to begin"
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "   You are in year 2525, and in charge of an anti-aircraft machine. Set the angle of the shot and destroy your target!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2760
      TabIndex        =   1
      Top             =   2040
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Set && Destroy®:  Cilck the begin button to start the game."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -2280
      Picture         =   "Set & Destroy.frx":1272
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   15360
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sami Eljabali
'Programming I
'Semester Project
Option Explicit

Private Sub cmdBegin_Click()
    'Takes you to the game itself
    frmGame.Show
    'This is done to leave only one form insight
    Unload Me
End Sub

Private Sub cmdBegin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If over the button(Dark Grey)
    cmdBegin.BackColor = &HC0C0C0
End Sub

Private Sub cmdDone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If over the button(Dark Grey)
    cmdDone.BackColor = &HC0C0C0
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If not over button(s)(Grey)
    If cmdBegin.BackColor = &HC0C0C0 Then
        cmdBegin.BackColor = &H8000000F
    End If
    If cmdDone.BackColor = &HC0C0C0 Then
        cmdDone.BackColor = &H8000000F
    End If
End Sub

Private Sub cmdDone_Click()
    'Takes you to credits
    'The "1" after ".show" makes the concentration only on the form itself*Got help
     frmEnding.Show 1
    'This is done to leave only one form insight
    Unload Me
End Sub
