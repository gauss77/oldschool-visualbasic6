VERSION 5.00
Begin VB.Form frmIntro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graphic Demo"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleMode       =   0  'User
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMovelbl2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   3000
   End
   Begin VB.Timer tmrColors 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   3000
   End
   Begin VB.Timer tmrMovelbl 
      Interval        =   10
      Left            =   0
      Top             =   3000
   End
   Begin VB.Image imgSmiley2 
      Height          =   1485
      Index           =   1
      Left            =   0
      Picture         =   "frmIntro.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image imgSmiley2 
      Height          =   1485
      Index           =   2
      Left            =   0
      Picture         =   "frmIntro.frx":0CB0
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image imgSmiley2 
      Height          =   1485
      Index           =   3
      Left            =   0
      Picture         =   "frmIntro.frx":1D14
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image imgSmiley 
      Height          =   1485
      Index           =   2
      Left            =   5760
      Picture         =   "frmIntro.frx":2D3B
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image imgSmiley 
      Height          =   1485
      Index           =   3
      Left            =   5760
      Picture         =   "frmIntro.frx":3D62
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image imgSmiley 
      Height          =   1485
      Index           =   1
      Left            =   5760
      Picture         =   "frmIntro.frx":4DC6
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblSami 
      Caption         =   "Sami's Graphic Demo!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label lblWelcome 
      Caption         =   "Welcome To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2700
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sami Eljabali
'Programming II
'Graphic Demo
Option Explicit

Private Sub Form_Load()
    AutoRedraw = True
    Randomize
    lblSami.Top = 2040
    lblSami.Left = 7080
    lblWelcome.Top = 1440
    lblWelcome.Left = -1440
End Sub

Private Sub tmrMovelbl_Timer()
    lblWelcome.Left = lblWelcome.Left + 15
    lblSami.Left = lblSami.Left - 17

        If (lblWelcome.Left = 2850) Then
            tmrMovelbl.Enabled = False
            tmrColors.Enabled = True
        End If
End Sub

Private Sub tmrColors_Timer()
    Dim intQbFunc As Integer
    Static intCounter As Integer
    Static intTemp As Integer
    
    If intCounter <> 15 Then
        
        intQbFunc = Int(Rnd * 10) 'Enabling Random color function
        
        Me.BackColor = QBColor(intQbFunc) 'Assigning color function
        lblSami.BackColor = QBColor(intQbFunc) 'Assigning color function
        lblWelcome.BackColor = QBColor(intQbFunc) 'Assigning color function
        intCounter = intCounter + 1
        intTemp = intTemp + 1
            'Enabling Smiley Faces
            If intTemp > 3 Then intTemp = 1
            imgSmiley(intTemp).Visible = True
            imgSmiley2(intTemp).Visible = True
            If intTemp > 1 Then imgSmiley(intTemp - 1).Visible = False
            If intTemp > 1 Then imgSmiley2(intTemp - 1).Visible = False
    Else 'If 15 colors have been displayed then...
        
        'Put the colors back to normal
        Me.BackColor = &H8000000F
        lblSami.BackColor = &H8000000F
        lblWelcome.BackColor = &H8000000F
                
        Dim intAR As Integer
            For intAR = 1 To 3
                imgSmiley(intAR).Visible = False 'Disable Smiley Faces
                imgSmiley2(intAR).Visible = False 'Disable Smiley Faces
            Next intAR
        tmrMovelbl2.Enabled = True 'Enables another timer
        tmrColors.Enabled = False 'Disables this timer
    End If
End Sub

Private Sub tmrMovelbl2_Timer()
    lblWelcome.Left = lblWelcome.Left + 15
    lblSami.Left = lblSami.Left - 17
        'Move the labels until one of them exit the screen
        If lblSami.Left = -3360 Or lblWelcome.Left = 6960 Then
            tmrMovelbl2.Enabled = False
            Unload Me
            frmChoice.Show
        End If
End Sub
