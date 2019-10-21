VERSION 5.00
Begin VB.Form frmStarfield 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graphic Demo"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrStarMove 
      Interval        =   5000
      Left            =   1440
      Top             =   6120
   End
   Begin VB.Timer tmrStar3 
      Interval        =   96
      Left            =   960
      Top             =   6120
   End
   Begin VB.Timer tmrStar2 
      Interval        =   64
      Left            =   480
      Top             =   6120
   End
   Begin VB.Timer tmrStar 
      Interval        =   128
      Left            =   0
      Top             =   6120
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   30
      X1              =   3480
      X2              =   3495
      Y1              =   1920
      Y2              =   1935
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   29
      X1              =   5640
      X2              =   5655
      Y1              =   5280
      Y2              =   5295
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   28
      X1              =   7080
      X2              =   7095
      Y1              =   6000
      Y2              =   6015
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   27
      X1              =   3000
      X2              =   3015
      Y1              =   6240
      Y2              =   6255
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   26
      X1              =   7200
      X2              =   7215
      Y1              =   1440
      Y2              =   1455
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   25
      X1              =   840
      X2              =   855
      Y1              =   2880
      Y2              =   2895
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   24
      X1              =   7320
      X2              =   7335
      Y1              =   4320
      Y2              =   4335
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   23
      X1              =   3960
      X2              =   3975
      Y1              =   720
      Y2              =   735
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   22
      X1              =   3120
      X2              =   3135
      Y1              =   3840
      Y2              =   3855
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   21
      X1              =   240
      X2              =   255
      Y1              =   720
      Y2              =   735
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   20
      X1              =   120
      X2              =   135
      Y1              =   3000
      Y2              =   3015
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   19
      X1              =   1680
      X2              =   1695
      Y1              =   6120
      Y2              =   6135
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   18
      X1              =   5400
      X2              =   5415
      Y1              =   4320
      Y2              =   4335
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   17
      X1              =   6000
      X2              =   6015
      Y1              =   2520
      Y2              =   2535
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   16
      X1              =   6000
      X2              =   6015
      Y1              =   2520
      Y2              =   2535
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   15
      X1              =   6000
      X2              =   6015
      Y1              =   2520
      Y2              =   2535
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   14
      X1              =   6000
      X2              =   6015
      Y1              =   2520
      Y2              =   2535
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   13
      X1              =   6000
      X2              =   6015
      Y1              =   2520
      Y2              =   2535
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   12
      X1              =   6000
      X2              =   6015
      Y1              =   2520
      Y2              =   2535
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   11
      X1              =   6000
      X2              =   6015
      Y1              =   2520
      Y2              =   2535
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   10
      X1              =   6000
      X2              =   6015
      Y1              =   2520
      Y2              =   2535
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   9
      X1              =   6000
      X2              =   6015
      Y1              =   2520
      Y2              =   2535
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   8
      X1              =   2640
      X2              =   2655
      Y1              =   2640
      Y2              =   2655
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   7
      X1              =   6000
      X2              =   6015
      Y1              =   3960
      Y2              =   3975
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   6
      X1              =   1320
      X2              =   1335
      Y1              =   5400
      Y2              =   5415
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   5
      X1              =   1080
      X2              =   1095
      Y1              =   480
      Y2              =   495
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   4
      X1              =   6240
      X2              =   6255
      Y1              =   600
      Y2              =   615
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   3
      X1              =   3960
      X2              =   3975
      Y1              =   5520
      Y2              =   5535
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   2
      X1              =   3840
      X2              =   3855
      Y1              =   240
      Y2              =   255
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   3960
      X2              =   3975
      Y1              =   3480
      Y2              =   3495
   End
   Begin VB.Line linStar 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   3240
      X2              =   3255
      Y1              =   1560
      Y2              =   1575
   End
End
Attribute VB_Name = "frmStarfield"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sami Eljabali
'Programming II
'Graphic Demo
Option Explicit

Private Sub Form_Click()
    'Exit if you click the screen
    Unload Me
End Sub

Private Sub Form_Load()
        
    Dim intRndX As Integer
    Dim intRndY As Integer
    Dim intIndex As Integer
    
    Me.AutoRedraw = True
    Me.BackColor = vbBlack
    Randomize
    
'Put stars in random locations at form load
For intIndex = 0 To 30
    intRndY = (6120 * Rnd + 1)
    intRndX = (7560 * Rnd + 1)
    linStar(intIndex).X1 = intRndX
    linStar(intIndex).X2 = intRndX
    linStar(intIndex).Y1 = intRndY
    linStar(intIndex).Y2 = intRndY
Next intIndex
               
End Sub

Private Sub tmrStar_Timer()
Dim intIndex As Integer

    'if they over size then put them back to normal size
    If linStar(1).BorderWidth = 13 Then
        For intIndex = 0 To 8
            linStar(intIndex).BorderWidth = 1
        Next intIndex
    End If

'9 of the 31 stars change its size in differnt timer intervals
For intIndex = 0 To 8
    linStar(intIndex).BorderWidth = linStar(intIndex).BorderWidth + 1
Next intIndex
End Sub

Private Sub tmrStar2_Timer()

Dim intIndex As Integer
    'if they over size then put them back to normal size
    If linStar(9).BorderWidth = 13 Then
        For intIndex = 9 To 20
            linStar(intIndex).BorderWidth = 1
        Next intIndex
    End If
'11 of the 31 stars change its size in differnt timer intervals
For intIndex = 9 To 20
    linStar(intIndex).BorderWidth = linStar(intIndex).BorderWidth + 1
Next intIndex
End Sub

Private Sub tmrStar3_Timer()

Dim intIndex As Integer
    'if they over size then put them back to normal size
    If linStar(21).BorderWidth = 13 Then
        For intIndex = 21 To 30
            linStar(intIndex).BorderWidth = 1
        Next intIndex
    End If
'9 of the 31 stars change its size in differnt timer intervals
For intIndex = 21 To 30
    linStar(intIndex).BorderWidth = linStar(intIndex).BorderWidth + 1
Next intIndex

End Sub

Private Sub tmrStarMove_Timer()

Dim intRndX As Integer
Dim intRndY As Integer
Dim intIndex As Integer

'Every 5 seconds change every star's location
For intIndex = 0 To 30
    intRndY = (6120 * Rnd + 1)
    intRndX = (7560 * Rnd + 1)
    linStar(intIndex).X1 = intRndX
    linStar(intIndex).X2 = intRndX
    linStar(intIndex).Y1 = intRndY
    linStar(intIndex).Y2 = intRndY
Next intIndex
End Sub
