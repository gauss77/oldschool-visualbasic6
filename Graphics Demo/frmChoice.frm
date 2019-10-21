VERSION 5.00
Begin VB.Form frmChoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Your Program."
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleMode       =   0  'User
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrLetterMover 
      Interval        =   10
      Left            =   1320
      Top             =   960
   End
   Begin VB.Timer tmrCountDown 
      Interval        =   600
      Left            =   2160
      Top             =   960
   End
   Begin VB.Timer tmrColorchanger 
      Interval        =   100
      Left            =   1680
      Top             =   960
   End
   Begin VB.Label lbl1 
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   5400
      TabIndex        =   9
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lbl1 
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   5280
      TabIndex        =   8
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lbl1 
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   5040
      TabIndex        =   7
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lbl1 
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   4920
      TabIndex        =   6
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lbl1 
      Caption         =   "f"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4800
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lbl1 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   4680
      TabIndex        =   4
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lbl1 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4440
      TabIndex        =   3
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lbl1 
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4320
      TabIndex        =   2
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lbl1 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3960
      TabIndex        =   1
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblChoose 
      Caption         =   "Entering the              !!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   4815
   End
End
Attribute VB_Name = "frmChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sami Eljabali
'Programming II
'Graphic Demo
Option Explicit
Private intCounter As Integer

Private Sub Form_Load()
    intCounter = 0
    Randomize
    Me.AutoRedraw = True
End Sub

Private Sub tmrColorchanger_Timer()
Dim intQbFunc As Integer
Dim intArray As Integer

    intQbFunc = Int(Rnd * 10)
    Me.BackColor = QBColor(intQbFunc)
    lblChoose.BackColor = QBColor(intQbFunc)
    For intArray = 0 To 8
        lbl1(intArray).BackColor = QBColor(intQbFunc)
    Next intArray
End Sub

Private Sub tmrCountDown_Timer()
   
    intCounter = intCounter + 1
    If intCounter = 5 Then
        Unload Me
        frmStarfield.Show
    End If
End Sub

Private Sub tmrLetterMover_Timer()
   If intCounter > 1 Then
    lbl1(0).Top = lbl1(0).Top - 10
    lbl1(0).Left = lbl1(0).Left - 10
    
    lbl1(1).Left = lbl1(1).Left - 10
    
    lbl1(2).Top = lbl1(2).Top + 6
    lbl1(2).Left = lbl1(2).Left - 10
    
    lbl1(3).Top = lbl1(3).Top + 10
    lbl1(3).Left = lbl1(3).Left - 10
    
    lbl1(4).Top = lbl1(4).Top - 10
    lbl1(4).Left = lbl1(4).Left + 10
    
    lbl1(5).Top = lbl1(5).Top - 6
    lbl1(5).Left = lbl1(5).Left + 10
    
    lbl1(6).Left = lbl1(6).Left + 10
    
    lbl1(7).Top = lbl1(7).Top + 10
    lbl1(7).Left = lbl1(7).Left + 10

    lbl1(8).Top = lbl1(8).Top + 5
    lbl1(8).Left = lbl1(8).Left + 10
   End If
End Sub
