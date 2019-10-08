VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmGame 
   Caption         =   "The Animal Game"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDone 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7440
      TabIndex        =   20
      Top             =   4680
      Width           =   1215
   End
   Begin MCI.MMControl MMDino 
      Height          =   330
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   "AudioWave"
      FileName        =   "A:\JurassicPark-Tyrannosaurus_rex-Roaring.wav"
   End
   Begin MCI.MMControl MMFrog 
      Height          =   330
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   "AudioWave"
      FileName        =   "A:\Frog.wav"
   End
   Begin VB.CommandButton cmdAnimal6 
      Caption         =   "Animal 6"
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdAnimal5 
      Caption         =   "Animal 5"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdAnimal4 
      Caption         =   "Animal 4"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdAnimal3 
      Caption         =   "Animal 3"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdAnimal2 
      Caption         =   "Animal 2"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin MCI.MMControl MMDuck 
      Height          =   330
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   "AudioWave"
      FileName        =   "A:\duck.wav"
   End
   Begin MCI.MMControl MMChicken 
      Height          =   330
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   "AudioWave"
      FileName        =   "A:\Chicken.wav"
   End
   Begin MCI.MMControl MMCat 
      Height          =   330
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   "AudioWave"
      FileName        =   "A:\Cat.wav"
   End
   Begin MCI.MMControl MMDog 
      Height          =   330
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   "AudioWave"
      FileName        =   "A:\dog.wav"
   End
   Begin VB.OptionButton optFrog 
      Caption         =   "FROG"
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Check"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   3960
      Width           =   1335
   End
   Begin VB.OptionButton optChicken 
      Caption         =   "CHICKEN"
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   3240
      Width           =   1095
   End
   Begin VB.OptionButton optDuck 
      Caption         =   "DUCK"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.OptionButton optCat 
      Caption         =   "CAT"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   3240
      Width           =   615
   End
   Begin VB.OptionButton optDog 
      Caption         =   "DOG"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose Animal"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdAnimal1 
      Caption         =   "Animal 1"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Frame 
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   3120
      Width           =   8055
   End
   Begin VB.Image imgCat 
      Height          =   1380
      Left            =   3120
      Picture         =   "Game.frx":0000
      Top             =   1680
      Width           =   1305
   End
   Begin VB.Image imgDog 
      Height          =   2220
      Left            =   120
      Picture         =   "Game.frx":0B61
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image imgDuck 
      Height          =   1500
      Left            =   1320
      Picture         =   "Game.frx":1D32
      Top             =   1560
      Width           =   2250
   End
   Begin VB.Image imgFrog 
      Height          =   1695
      Left            =   6360
      Picture         =   "Game.frx":2EAC
      Top             =   1440
      Width           =   2250
   End
   Begin VB.Image imgChicken 
      Height          =   2250
      Left            =   4440
      Picture         =   "Game.frx":43EA
      Top             =   1320
      Width           =   1830
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
Private strAnswer As String
Private intTries As Integer

Private Sub cmdAnimal1_Click()
   'Playing Sounds
    MMDog.Command = "play"
    MMDog.Command = "prev"
    cmdAnimal2.Visible = False
    cmdAnimal3.Visible = False
    cmdAnimal4.Visible = False
    cmdAnimal5.Visible = False
    cmdAnimal6.Visible = False
    strAnswer = "Dog"
End Sub

Private Sub cmdAnimal2_Click()
   'Playing Sounds
    MMCat.Command = "play"
    MMCat.Command = "prev"
    cmdAnimal1.Visible = False
    cmdAnimal3.Visible = False
    cmdAnimal4.Visible = False
    cmdAnimal5.Visible = False
    cmdAnimal6.Visible = False
    strAnswer = "Cat"
End Sub

Private Sub cmdAnimal3_Click()
   'Playing Sounds
    MMChicken.Command = "play"
    MMChicken.Command = "prev"
    cmdAnimal1.Visible = False
    cmdAnimal2.Visible = False
    cmdAnimal4.Visible = False
    cmdAnimal5.Visible = False
    cmdAnimal6.Visible = False
    strAnswer = "Chicken"
End Sub

Private Sub cmdAnimal4_Click()
   'Playing Sounds
    MMDuck.Command = "play"
    MMDuck.Command = "prev"
    cmdAnimal1.Visible = False
    cmdAnimal2.Visible = False
    cmdAnimal3.Visible = False
    cmdAnimal5.Visible = False
    cmdAnimal6.Visible = False
    strAnswer = "Duck"
End Sub

Private Sub cmdAnimal5_Click()
   'Playing Sounds
    MMFrog.Command = "play"
    MMFrog.Command = "prev"
    cmdAnimal1.Visible = False
    cmdAnimal2.Visible = False
    cmdAnimal3.Visible = False
    cmdAnimal4.Visible = False
    cmdAnimal6.Visible = False
    strAnswer = "Frog"
End Sub

Private Sub cmdAnimal6_Click()
   'Playing Sounds
    MMDino.Command = "play"
    MMDino.Command = "prev"
    cmdAnimal1.Visible = False
    cmdAnimal2.Visible = False
    cmdAnimal3.Visible = False
    cmdAnimal4.Visible = False
    cmdAnimal5.Visible = False
    strAnswer = "Dino"
End Sub

Private Sub cmdChoose_Click()
    'Enabling pics and option buttons
    ImgCat.Visible = True
    ImgDog.Visible = True
    ImgChicken.Visible = True
    ImgFrog.Visible = True
    ImgDuck.Visible = True
    optCat.Visible = True
    optDog.Visible = True
    optFrog.Visible = True
    optChicken.Visible = True
    optDuck.Visible = True
    Frame.Visible = True
    cmdOK.Visible = True
End Sub

Private Sub cmdOK_Click()
    'Checking answer
    If strAnswer = "Dog" And optDog.Value = True Then
        MsgBox "You are correct!!"
        MsgBox "It took you" & " " & intTries & " tries."
        Unload Me 'Closing game
        frmQuestion.Show 'Openning Questions
        
    ElseIf strAnswer = "Cat" And optCat.Value = True Then
        MsgBox "You are correct!!"
        MsgBox "It took you " & intTries & " tries."
        Unload Me 'Closing game
        frmQuestion.Show 'Openning Questions
     
     ElseIf strAnswer = "Chicken" And optChicken.Value = True Then
        MsgBox "You are correct!!"
        MsgBox "It took you" & " " & intTries & " tries."
        Unload Me 'Closing game
        frmQuestion.Show 'Openning Questions
    
    ElseIf strAnswer = "Duck" And optDuck.Value = True Then
        MsgBox "You are correct!!"
        MsgBox "It took you" & " " & intTries & " tries."
        Unload Me 'Closing game
        frmQuestion.Show 'Openning Questions
        
    ElseIf strAnswer = "Frog" And optFrog.Value = True Then
        MsgBox "You are correct!!"
        MsgBox "It took you" & " " & intTries & " tries."
        Unload Me 'Closing game
        frmQuestion.Show 'Openning Questions
       
    Else
        MsgBox "You are wrong!!"
        intTries = intTries + 1
    End If
End Sub

Private Sub Form_Load()
    'Disabling the view of unwanted objects
    ImgCat.Visible = False
    ImgDog.Visible = False
    ImgChicken.Visible = False
    ImgFrog.Visible = False
    ImgDuck.Visible = False
    optCat.Visible = False
    optDog.Visible = False
    optFrog.Visible = False
    optChicken.Visible = False
    optDuck.Visible = False
    Frame.Visible = False
    cmdOK.Visible = False
    MMDog.Visible = False
    MMCat.Visible = False
    MMChicken.Visible = False
    MMDuck.Visible = False
    MMFrog.Visible = False
    MMDino.Visible = False

    'Enabling Sound of Dog
    MMDog.DeviceType = "WaveAudio"
    MMDog.FileName = "dog.wav"
    MMDog.Command = "open"

    'Enabling Sound of Cat
    MMCat.DeviceType = "WaveAudio"
    MMCat.FileName = "Cat.wav"
    MMCat.Command = "open"
    
    'Enabling Sound of Chicken
    MMChicken.DeviceType = "WaveAudio"
    MMChicken.FileName = "chicken.wav"
    MMChicken.Command = "open"
    
    'Enabling Sound of Duck
    MMDuck.DeviceType = "WaveAudio"
    MMDuck.FileName = "duck.wav"
    MMDuck.Command = "open"
    
    'Enabling Sound of Frog
    MMFrog.DeviceType = "WaveAudio"
    MMFrog.FileName = "Frog.wav"
    MMFrog.Command = "open"
    
    'Enabling Sound of Dino
    MMDino.DeviceType = "WaveAudio"
    MMDino.FileName = "JurassicPark-Tyrannosaurus_rex-Roaring.wav"
    MMDino.Command = "open"
End Sub

Private Sub cmdDone_Click()
Dim strExit As String
    Do While strExit <> "yes" Or strExit <> "no"
        strExit = InputBox("Are you sure you want to leave? (yes or no)", "Animal Game")
        If strExit = "yes" Then
            Unload Me
            Exit Sub
        ElseIf strExit = "no" Then
            Unload Me
            frmGame.Show
            Exit Sub
        End If
    Loop
End Sub
