VERSION 5.00
Begin VB.Form frmCos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cos & Sine Waves."
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9570
   Icon            =   "frmCos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSine 
      Caption         =   "Sine"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Cos"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Default         =   -1  'True
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Done"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   6
      Left            =   0
      Top             =   8040
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7155
      ScaleWidth      =   9315
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.Line Line2 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   7200
      End
      Begin VB.Shape shp2 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         Shape           =   3  'Circle
         Top             =   3600
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   -360
         X2              =   9360
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Shape shp 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         Shape           =   3  'Circle
         Top             =   3600
         Width           =   375
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Blue Ball = Sine Wave"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Red Ball = Cos Wave"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   7440
      Width           =   1695
   End
End
Attribute VB_Name = "frmCos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sami Eljabali
'Cos & Sine Waves
Option Explicit

Private Sub Picture1_Click()
    Picture1.AutoRedraw = True
End Sub

Private Sub Timer1_Timer()
    Dim Y As Double
    Dim Y2 As Double
    
    Y = Cos(50 * shp.Left / 9375)
    Y2 = Sin(50 * shp2.Left / 9375)
    shp.Left = shp.Left + 20
        If shp.Left >= 9300 Then shp.Left = 0
        If shp2.Left >= 9300 Then shp2.Left = 0
    shp.Top = 3600 + Y * 1000
    
    'Y2 = Sin(50 * shp2.Left / 9375)
    shp2.Left = shp2.Left + 20
    shp2.Top = 3600 + Y2 * 1000
End Sub

Private Sub cmdCos_Click()
    If shp.Visible = True Then
        shp.Visible = False
    Else
        shp.Visible = True
    End If
End Sub

Private Sub cmdSine_Click()
    If shp2.Visible = True Then
        shp2.Visible = False
    Else
        shp2.Visible = True
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
    frmCos.Show
End Sub


