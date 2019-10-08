VERSION 5.00
Begin VB.Form frmQuestion 
   Caption         =   "Questions"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAgain 
      Caption         =   "Play Again"
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Image ImgFrog 
      Height          =   1695
      Left            =   3960
      Picture         =   "frmQuestion.frx":0000
      Top             =   3360
      Width           =   2250
   End
   Begin VB.Image ImgDuck 
      Height          =   1500
      Left            =   360
      Picture         =   "frmQuestion.frx":153E
      Top             =   3240
      Width           =   2250
   End
   Begin VB.Image ImgDog 
      Height          =   2220
      Left            =   6720
      Picture         =   "frmQuestion.frx":26B8
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image ImgChicken 
      Height          =   2250
      Left            =   2280
      Picture         =   "frmQuestion.frx":3889
      Top             =   480
      Width           =   1830
   End
   Begin VB.Image ImgCat 
      Height          =   1380
      Left            =   240
      Picture         =   "frmQuestion.frx":4AF8
      Top             =   960
      Width           =   1305
   End
End
Attribute VB_Name = "frmQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sami Eljabali
'Programming I
'Semester Project
Option Explicit

Private Sub ImgCat_Click()
    Dim strTextEntered3 As String
        strTextEntered3 = InputBox("What is a baby Cat called?", "Cat")
    If strTextEntered3 = "" Then
        MsgBox "Please Type in your answer"
    ElseIf strTextEntered3 = "kitten" Then
        MsgBox " Congratualtions, You are Correct"
    ElseIf strTextEntered3 <> "kitten" Then
        MsgBox "Sorry, wrong answer, please try again"
    End If
End Sub

Private Sub ImgChicken_Click()
    Dim strTextEntered2 As String
        strTextEntered2 = InputBox("What is a baby Chicken called?", "Chicken")
    If strTextEntered2 = "" Then
        MsgBox "Please Type in your answer"
    ElseIf strTextEntered2 = "chick" Then
        MsgBox " Congratualtions, You are Correct"
    ElseIf strTextEntered2 <> "chick" Then
        MsgBox "Sorry, wrong answer, please try again"
    End If
End Sub

Private Sub imgDog_Click()
    Dim strTextEntered4 As String
        strTextEntered4 = InputBox("What is a baby Dog called?", "Dog")
    If strTextEntered4 = "" Then
        MsgBox "Please Type in your answer"
    ElseIf strTextEntered4 = "puppy" Then
        MsgBox " Congratualtions, You are Correct"
    ElseIf strTextEntered4 <> "puppy" Then
        MsgBox "Sorry, wrong answer, please try again"
    End If
End Sub

Private Sub imgDuck_Click()
    Dim strTextEntered As String
        strTextEntered = InputBox("What is a baby Duck called?", "Duck")
    If strTextEntered = "" Then
        MsgBox "Please Type in your answer"
    ElseIf strTextEntered = "duckling" Then
        MsgBox " Congratualtions, You are Correct"
    ElseIf strTextEntered <> "duckling" Then
        MsgBox "Sorry, wrong answer, please try again"
    End If
End Sub

Private Sub imgFrog_Click()
    Dim strTextEntered1 As String
        strTextEntered1 = InputBox("What is a baby Frog called?", "Frog")
    If strTextEntered1 = "" Then
        MsgBox "Please Type in your answer"
    ElseIf strTextEntered1 = "tadpole" Then
        MsgBox " Congratualtions, You are Correct"
    ElseIf strTextEntered1 <> "tadpole" Then
        MsgBox "Sorry, wrong answer, please try again"
    End If
End Sub

Private Sub cmdAgain_Click()
    Unload Me
    frmGame.Show
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

