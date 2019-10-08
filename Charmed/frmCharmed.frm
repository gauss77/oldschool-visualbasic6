VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmCharmed 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Charmed"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox imgCharmed8 
      Height          =   3375
      Left            =   -840
      Picture         =   "frmCharmed.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   6195
      TabIndex        =   15
      Top             =   -1080
      Width           =   6255
   End
   Begin VB.PictureBox imgCharmed7 
      Height          =   3375
      Left            =   -840
      Picture         =   "frmCharmed.frx":5992
      ScaleHeight     =   3315
      ScaleWidth      =   6195
      TabIndex        =   14
      Top             =   -1080
      Width           =   6255
   End
   Begin VB.PictureBox imgCharmed6 
      Height          =   3375
      Left            =   -840
      Picture         =   "frmCharmed.frx":B31A
      ScaleHeight     =   3315
      ScaleWidth      =   6195
      TabIndex        =   13
      Top             =   -1080
      Width           =   6255
   End
   Begin VB.PictureBox imgCharmed5 
      Height          =   3375
      Left            =   -840
      Picture         =   "frmCharmed.frx":10C69
      ScaleHeight     =   3315
      ScaleWidth      =   6195
      TabIndex        =   12
      Top             =   -1080
      Width           =   6255
   End
   Begin VB.PictureBox imgCharmed4 
      Height          =   3375
      Left            =   -840
      Picture         =   "frmCharmed.frx":165AE
      ScaleHeight     =   3315
      ScaleWidth      =   6195
      TabIndex        =   11
      Top             =   -1080
      Width           =   6255
   End
   Begin VB.PictureBox imgCharmed3 
      Height          =   3375
      Left            =   -840
      Picture         =   "frmCharmed.frx":8ED98
      ScaleHeight     =   3315
      ScaleWidth      =   6195
      TabIndex        =   10
      Top             =   -1080
      Width           =   6255
   End
   Begin VB.PictureBox imgCharmed2 
      Height          =   3375
      Left            =   -840
      Picture         =   "frmCharmed.frx":946B2
      ScaleHeight     =   3315
      ScaleWidth      =   6195
      TabIndex        =   9
      Top             =   -1080
      Width           =   6255
   End
   Begin VB.PictureBox imgCharmed 
      Height          =   3375
      Left            =   -840
      Picture         =   "frmCharmed.frx":99FF2
      ScaleHeight     =   3315
      ScaleWidth      =   6195
      TabIndex        =   8
      Top             =   -1080
      Width           =   6255
   End
   Begin VB.Timer tmrEnd 
      Interval        =   1
      Left            =   3240
      Top             =   0
   End
   Begin VB.Timer tmrFibi 
      Interval        =   1
      Left            =   2760
      Top             =   0
   End
   Begin VB.Timer tmrPiper 
      Interval        =   1
      Left            =   2280
      Top             =   0
   End
   Begin VB.Timer tmrChanging 
      Interval        =   80
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrCharmed 
      Interval        =   1
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer tmrC 
      Interval        =   100
      Left            =   840
      Top             =   1560
   End
   Begin VB.Timer tmrH 
      Interval        =   100
      Left            =   1440
      Top             =   1560
   End
   Begin VB.Timer tmrA 
      Interval        =   100
      Left            =   2040
      Top             =   1560
   End
   Begin VB.Timer tmrR 
      Interval        =   100
      Left            =   2640
      Top             =   1560
   End
   Begin VB.Timer tmrM 
      Interval        =   100
      Left            =   3240
      Top             =   1560
   End
   Begin VB.Timer tmrE 
      Interval        =   100
      Left            =   3840
      Top             =   1560
   End
   Begin VB.Timer tmrD 
      Interval        =   100
      Left            =   4440
      Top             =   1560
   End
   Begin VB.Timer tmrColor 
      Interval        =   200
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tmrMover 
      Interval        =   2
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer tmrPru 
      Interval        =   1
      Left            =   1800
      Top             =   0
   End
   Begin MCI.MMControl MMCharmed 
      Height          =   330
      Left            =   -840
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      PlayEnabled     =   -1  'True
      DeviceType      =   ".wav"
      FileName        =   "D:\VB6\Programming II\Charmed\T.A.T.U - How Soon Is Now-!.wav"
   End
   Begin VB.Label lblD 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.Image imgFibi 
      Height          =   1890
      Left            =   3720
      Picture         =   "frmCharmed.frx":A0DAE
      Top             =   240
      Width           =   1680
   End
   Begin VB.Image imgPiper 
      Height          =   1890
      Left            =   3720
      Picture         =   "frmCharmed.frx":A23BC
      Top             =   240
      Width           =   1680
   End
   Begin VB.Image imgPru 
      Height          =   1890
      Left            =   3720
      Picture         =   "frmCharmed.frx":A33D9
      Top             =   240
      Width           =   1680
   End
   Begin VB.Label lblC 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblH 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblA 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblR 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblM 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblE 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.Image imgLogo 
      Height          =   2325
      Left            =   0
      Picture         =   "frmCharmed.frx":A48B6
      Top             =   0
      Width           =   5475
   End
End
Attribute VB_Name = "frmCharmed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sami Eljabali
'Programming II
'Charmed
Option Explicit
Private intCount As Integer
Private intTemp As Integer
Private intTemp2 As Integer
'FORM LOAD
Private Sub Form_Load()
    'SIZES
    Me.Height = 2745
    Me.Width = 5490
    Me.ScaleHeight = 2265
    Me.ScaleWidth = 5400
    'MUSIC
    MMCharmed.DeviceType = "WaveAudio"
    MMCharmed.FileName = "T.A.T.U - How Soon Is Now-!.wav"
    MMCharmed.Command = "Open"
    MMCharmed.Command = "Play"
    MMCharmed.Command = "Prev"
    MMCharmed.PlayEnabled = True
    'OTHERS
    Randomize
    Me.AutoRedraw = True
    intCount = 0
    intTemp = 0
    intTemp2 = 0
    'PICTURES
    imgPru.Visible = False
    imgPiper.Visible = False
    imgFibi.Visible = False
    imgLogo.Visible = False
    imgCharmed.Visible = False
    imgCharmed2.Visible = False
    imgCharmed3.Visible = False
    imgCharmed4.Visible = False
    imgCharmed5.Visible = False
    imgCharmed6.Visible = False
    imgCharmed7.Visible = False
    imgCharmed8.Visible = False
End Sub

Private Sub tmrChanging_Timer()
'CONSTANT COUNTER
    intCount = intCount + 1
End Sub

Private Sub tmrColor_Timer()
'BACKGROUND COLOR CHANGER
    If Not intCount >= 62 Then
        Me.BackColor = QBColor(Int((7 - 1 + 1) * Rnd + 1))
    Else
        Me.BackColor = &H8000000F
        tmrColor.Enabled = False
    End If
    End Sub
'LETTER CHANGER
Private Sub tmrC_Timer()
    lblC = Chr(Int((89 - 65 + 1) * Rnd + 65))
End Sub

Private Sub tmrH_Timer()
    lblH = Chr(Int((89 - 65 + 1) * Rnd + 65))
End Sub

Private Sub tmrA_Timer()
    lblA = Chr(Int((89 - 65 + 1) * Rnd + 65))
End Sub

Private Sub tmrR_Timer()
    lblR = Chr(Int((89 - 65 + 1) * Rnd + 65))
End Sub

Private Sub tmrM_Timer()
    lblM = Chr(Int((89 - 65 + 1) * Rnd + 65))
End Sub

Private Sub tmrE_Timer()
    lblE = Chr(Int((89 - 65 + 1) * Rnd + 65))
End Sub

Private Sub tmrD_Timer()
    lblD = Chr(Int((89 - 65 + 1) * Rnd + 65))
End Sub

Private Sub tmrCharmed_Timer()
'CHARMED
    If intCount = 30 Then
        tmrC.Enabled = False
        lblC = "C"
    End If
    
    If intCount = 50 Then
        tmrH.Enabled = False
        lblH = "H"
    End If
    If intCount = 55 Then
        tmrA.Enabled = False
        lblA = "A"
    End If
    If intCount = 62 Then
        tmrR.Enabled = False
        lblR = "R"
    End If
    If intCount = 62 Then
        tmrM.Enabled = False
        lblM = "M"
        imgLogo.Visible = True
    End If
    If intCount = 60 Then
        tmrE.Enabled = False
        lblE = "E"
    End If
    If intCount = 40 Then
        tmrD.Enabled = False
        lblD = "D"
    End If
    
    If intCount = 64 Then
        tmrCharmed.Enabled = False
    End If
End Sub

Private Sub tmrMover_Timer()
'MOVING LETTERS
If intCount >= 80 And intCount < 95 Then
    lblC.Top = lblC.Top - 10
    lblH.Top = lblH.Top + 10
    lblA.Top = lblA.Top - 10
    lblR.Top = lblR.Top + 10
    lblM.Top = lblM.Top - 10
    lblE.Top = lblE.Top + 10
    lblD.Top = lblD.Top - 10
End If
    
If intCount >= 95 Then
    lblC.Top = lblC.Top + 10
    lblH.Top = lblH.Top - 10
    lblA.Top = lblA.Top + 10
    lblR.Top = lblR.Top - 10
    lblM.Top = lblM.Top + 10
    lblE.Top = lblE.Top - 10
    lblD.Top = lblD.Top + 10
End If
    
If intCount >= 120 Then
    Do While Not lblC.Top = 960
        lblC.Top = lblC.Top - 1
        lblH.Top = lblH.Top + 1
        lblA.Top = lblA.Top - 1
        lblR.Top = lblR.Top + 1
        lblM.Top = lblM.Top - 1
        lblE.Top = lblE.Top + 1
        lblD.Top = lblD.Top - 1
    Loop
tmrMover.Enabled = False
End If
End Sub

Private Sub tmrPru_Timer()
'PRU
If tmrMover.Enabled = False Then
    lblC.Caption = "P"
    lblH.Caption = "r"
    lblA.Caption = "u"
    lblR.Caption = ""
    lblM.Caption = ""
    lblE.Caption = ""
    lblD.Caption = ""
    imgPru.Visible = True
End If
End Sub

Private Sub tmrPiper_Timer()
'PIPER
If intCount = 151 Then
    tmrPru.Enabled = False
    imgPru.Visible = False
        
    Do While lblC.Top <> 0 And intTemp = 0
        lblC.Top = lblC.Top - 1
        lblH.Top = lblH.Top + 1
        lblA.Top = lblA.Top - 1
        lblR.Top = lblR.Top + 1
        lblM.Top = lblM.Top - 1
    Loop
    intTemp = 1
    If lblC.Top = 0 Then
        Do While Not lblC.Top = 960
            lblC.Top = lblC.Top + 1
            lblH.Top = lblH.Top - 1
            lblA.Top = lblA.Top + 1
            lblR.Top = lblR.Top - 1
            lblM.Top = lblM.Top + 1
            lblE.Top = lblE.Top - 1
            lblD.Top = lblD.Top + 1
        Loop
        imgPiper.Visible = True
        lblC.Caption = "P"
        lblH.Caption = "i"
        lblA.Caption = "p"
        lblR.Caption = "e"
        lblM.Caption = "r"
    End If
End If
End Sub

Private Sub tmrFibi_Timer()
'Fibi
If intCount > 189 Then
    tmrPiper.Enabled = False
    imgPiper.Visible = False
    
    Do While lblC.Top <> 0 And intTemp2 = 0
        lblC.Top = lblC.Top - 1
        lblH.Top = lblH.Top + 1
        lblA.Top = lblA.Top - 1
        lblR.Top = lblR.Top + 1
        lblM.Top = lblM.Top - 1
    Loop
    intTemp2 = 1
    If lblC.Top = 0 Then
        Do While Not lblC.Top = 960
            lblC.Top = lblC.Top + 1
            lblH.Top = lblH.Top - 1
            lblA.Top = lblA.Top + 1
            lblR.Top = lblR.Top - 1
            lblM.Top = lblM.Top + 1
            lblE.Top = lblE.Top - 1
            lblD.Top = lblD.Top + 1
        Loop
        imgFibi.Visible = True
        lblC.Caption = "F"
        lblH.Caption = "i"
        lblA.Caption = "b"
        lblR.Caption = "i"
        lblM.Caption = ""
    End If
tmrFibi.Enabled = False
End If
End Sub
'END
Private Sub tmrEnd_Timer()
    If intCount > 229 And intCount < 260 Then
       If lblH.Top <> 0 Then
        lblC.Left = lblC.Left - 10
        lblH.Top = lblH.Top - 10
        lblA.Top = lblA.Top + 10
        lblR.Left = lblR.Left + 10
      ElseIf lblH.Top = 0 Then
        lblC = ""
        lblH = ""
        lblA = ""
        lblR = ""
        imgFibi.Visible = False
        imgLogo.Visible = False
        imgCharmed.Visible = True
      End If
    End If
    
    If intCount = 260 Then
        imgCharmed.Visible = False
        imgCharmed2.Visible = True
    End If
    
    If intCount = 263 Then
        imgCharmed2.Visible = False
        imgCharmed3.Visible = True
    End If
    
    If intCount = 266 Then
        imgCharmed3.Visible = False
        imgCharmed4.Visible = True
    End If
    
    If intCount = 269 Then
        imgCharmed4.Visible = False
        imgCharmed5.Visible = True
    End If
    
    If intCount = 272 Then
        imgCharmed5.Visible = False
        imgCharmed6.Visible = True
    End If
    
    If intCount = 275 Then
        imgCharmed6.Visible = False
        imgCharmed7.Visible = True
    End If
    
    If intCount = 278 Then
        imgCharmed7.Visible = False
        imgCharmed8.Visible = True
    End If
    
    If intCount = 281 Then
        imgCharmed8.Visible = False
        imgCharmed.Visible = True
    End If
    
    If intCount = 284 Then
        imgCharmed.Visible = False
        imgCharmed8.Visible = True
    End If
    
    If intCount = 287 Then
        imgCharmed8.Visible = False
        imgCharmed.Visible = True
    End If
    
    If intCount = 290 Then
        imgCharmed.Visible = False
        imgCharmed8.Visible = True
    End If
    
    If intCount = 293 Then
        imgCharmed8.Visible = False
        imgCharmed.Visible = True
    End If
    
    If intCount = 302 Then
        Unload Me
        frmCharmed.Show
    End If
End Sub
'SHOWIING THE LYRICS
Private Sub Form_Click()
    Call ShowLyrics
End Sub

Private Sub imgLogo_Click()
    Call ShowLyrics
End Sub

Private Sub imgCharmed_Click()
    Call ShowLyrics
End Sub

Private Sub imgFibi_Click()
    Call ShowLyrics
End Sub

Private Sub imgPiper_Click()
    Call ShowLyrics
End Sub

Private Sub imgPru_Click()
    Call ShowLyrics
End Sub

Private Sub lblH_Click()
    Call ShowLyrics
End Sub

Private Sub lblA_Click()
    Call ShowLyrics
End Sub

Private Sub lblR_Click()
    Call ShowLyrics
End Sub

Private Sub lblM_Click()
    Call ShowLyrics
End Sub

Private Sub lblE_Click()
    Call ShowLyrics
End Sub

Private Sub lblD_Click()
    Call ShowLyrics
End Sub

Public Sub ShowLyrics()
    frmLyrics.Show
End Sub
