VERSION 5.00
Begin VB.Form frmLyrics 
   Caption         =   """How Soon Is Now?"" by Tatu"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5340
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   5340
   Begin VB.TextBox txtLyrics 
      Height          =   7095
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmLyrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sami El Jabali
'Programming II
'Charmed
Option Explicit

Private Sub Form_Load()
Dim Lyrics As Integer
Dim strFile As String
Dim txtFile As String
    
    Me.Top = 0
    Me.Left = 0
    Lyrics = FreeFile
    strFile = txtLyrics
    Open "Charmed.txt" For Input As #Lyrics
        Do While Not EOF(Lyrics)
            Line Input #Lyrics, strFile
            txtFile = txtFile & strFile & vbCrLf
        Loop
        txtLyrics = txtFile
    Close #Lyrics
End Sub

Private Sub Form_Resize()
    txtLyrics.Width = Me.ScaleWidth
    txtLyrics.Height = Me.ScaleHeight
End Sub
