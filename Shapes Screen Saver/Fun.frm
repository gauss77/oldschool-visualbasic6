VERSION 5.00
Begin VB.Form frmFun 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4860
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6855
   ControlBox      =   0   'False
   Icon            =   "Fun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Fun.frx":030A
   Palette         =   "Fun.frx":0BD4
   ScaleHeight     =   4860
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrSquare 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   3480
      Top             =   2160
   End
   Begin VB.Timer tmrCircle 
      Interval        =   20
      Left            =   2880
      Top             =   2160
   End
End
Attribute VB_Name = "frmFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sami Eljabali
'Programming II
'Testing
Option Explicit
Dim intCounter As Integer
Dim intCounter2 As Integer

Private Sub Form_Load()
    Me.WindowState = vbMaximized
    Me.BackColor = vbBlack
    Me.AutoRedraw = True
    intCounter = 0
    intCounter2 = 0
End Sub

Private Sub tmrCircle_Timer()
    Dim Radius As Long
    Dim X As Long
    Dim Y As Long
    
    intCounter = intCounter + 1
    If intCounter = 1 Then
        Me.BackColor = vbWhite
        Me.BackColor = vbBlack
    End If

If intCounter < 200 Then
    X = Int(ScaleWidth * Rnd + 1)
    Y = Int(ScaleHeight * Rnd + 1)
    Radius = Int(1000 * Rnd + 700)
    Me.Circle (X, Y), Radius, vbWhite
Else
    tmrCircle.Enabled = False
    tmrSquare.Enabled = True
End If
End Sub

Private Sub tmrSquare_Timer()
   
   intCounter2 = intCounter2 + 1
    
    If intCounter2 = 1 Then
        Me.BackColor = vbWhite
        Me.BackColor = vbBlack
    End If
    
    Dim Radius As Long
    Dim X As Long
    Dim Y As Long
    Dim X2 As Long
    Dim Y2 As Long
    
    X = Int(ScaleWidth * Rnd + 1)
    X2 = Int(ScaleWidth * Rnd + 1)
    Y = Int(ScaleHeight * Rnd + 1)
    Y2 = Int(ScaleHeight * Rnd + 1)
 
If intCounter2 < 120 Then
    Me.Line (X, X2)-(Y, Y2), vbWhite, B
Else
    tmrCircle = True
    tmrSquare = False
    intCounter = 0
    intCounter2 = 0
End If
End Sub

Private Sub Form_Click()
    Unload Me
End Sub
