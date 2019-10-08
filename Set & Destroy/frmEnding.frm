VERSION 5.00
Begin VB.Form frmEnding 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "What Do You Think?!"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   Icon            =   "frmEnding.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEnding.frx":0442
   ScaleHeight     =   2580
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdBye 
      Cancel          =   -1  'True
      Caption         =   "Good-Bye "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Good-Bye!!"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Did you like it?!"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Image ImgHappy 
      Height          =   480
      Left            =   2880
      Picture         =   "frmEnding.frx":0784
      ToolTipText     =   "Ofcourse!!"
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "This game was done by Sami Eljabali. I hope you have enjoyed playing Set && Destroy®."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Image ImgSad 
      Height          =   480
      Left            =   2880
      Picture         =   "frmEnding.frx":0BC6
      ToolTipText     =   "No!!"
      Top             =   1440
      Width           =   480
   End
End
Attribute VB_Name = "frmEnding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sami Eljabali
'Programming I
'Semester Project
Option Explicit

Private Sub CmdBye_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If over the button(Dark Grey)
    CmdBye.BackColor = &HC0C0C0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If not over button(s)(Grey)
    If CmdBye.BackColor = &HC0C0C0 Then
        CmdBye.BackColor = &H8000000F
    End If
End Sub

Private Sub Form_Load()
    'Setting defult faces
    ImgSad.Visible = True
    ImgHappy.Visible = False
End Sub

Private Sub imghappy_Click()
    'If Happy make sad and vice versa
    If ImgSad.Visible = True Then
    ImgHappy.Visible = False
    Else
    ImgHappy.Visible = True
    End If
End Sub

Private Sub imgsad_Click()
    'If Happy make sad and vice versa
    If ImgHappy.Visible = True Then
    ImgHappy.Visible = False
    Else
    ImgHappy.Visible = True
    End If
End Sub

Private Sub CmdBye_Click()
    'If you liked the game or not
    If ImgHappy.Visible = True Then
        MsgBox "I'm glad that you liked it!!"
        End
    Else
        MsgBox "         What?"
        MsgBox "You are not happy?!"
        MsgBox "Fine then..."
        MsgBox "Virus Detected!!", vbCritical, "Alert!!"
        MsgBox "Happy now?!", , "Set & Destroy®"
        End
    End If
    'Ends the entire program
    End
End Sub
