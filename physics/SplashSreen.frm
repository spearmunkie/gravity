VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00404040&
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   16  'Merge Pen
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3360
      Top             =   2160
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "(Click anywhere to continue)"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   930
      Left            =   4680
      TabIndex        =   2
      Top             =   5760
      Width           =   8835
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Vader's Revenge!"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1890
      Left            =   2640
      TabIndex        =   1
      Top             =   4200
      Width           =   10605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to "
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1890
      Left            =   3960
      TabIndex        =   0
      Top             =   2880
      Width           =   7875
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   20
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   19
      Left            =   9120
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF80FF&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   18
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   6120
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   17
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   600
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   16
      Left            =   120
      Shape           =   3  'Circle
      Top             =   4560
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   15
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   6960
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   14
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00404080&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   13
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00008080&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   12
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   11
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   10
      Left            =   960
      Shape           =   3  'Circle
      Top             =   5760
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   9
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H80000001&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   8
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   7
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   6
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   5
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   4
      Left            =   9960
      Shape           =   3  'Circle
      Top             =   7200
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   3
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   2
      Left            =   720
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   1
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   480
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   0
      Left            =   720
      Shape           =   3  'Circle
      Top             =   480
      Width           =   750
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Circal
    X As Integer
    Y As Integer
       
    xIncrement As Integer
    yIncrement As Integer
End Type
Dim shapeInit(0 To 20) As Circal
Public difficulty As Integer
Public opt1 As Boolean
Public opt2 As Boolean
Public opt3 As Boolean
Public opt4 As Boolean
Public Colour As Long

Private Function Collide(x1 As Integer, y1 As Integer, x2 As Integer, _
                          y2 As Integer, rect As Integer) As Boolean
    If Abs(x1 - x2) <= rect And Abs(y1 - y2) <= rect Then
        Collide = True
    Else
        Collide = False
    End If
End Function
    
Private Sub Form_Load()
    Dim X As Integer
    For X = 0 To 20
       shapeInit(X).X = Shape1(X).Left + 750
       shapeInit(X).Y = Shape1(X).Top + 750
         
       shapeInit(X).xIncrement = 750 / 10
       shapeInit(X).yIncrement = 750 / 10
    Next X
   difficulty = 1
   opt1 = False
   opt2 = False
   opt3 = False
   opt4 = False
   Colour = vbBlack
End Sub


Private Sub Label_Click(Index As Integer)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button > 0 Then
      frmSplash.Hide
      frmMain.Show
   End If
End Sub
Private Sub Label1_Click()
   frmSplash.Hide
   frmMain.Show
End Sub
Private Sub Label2_Click()
   frmSplash.Hide
   frmMain.Show
End Sub
Private Sub Label3_Click()
   frmSplash.Hide
   frmMain.Show
End Sub
Private Sub Label4_Click()
   frmSplash.Hide
   frmMain.Show
End Sub
Private Sub Label5_Click()
   frmSplash.Hide
   frmMain.Show
End Sub

Private Sub Timer1_Timer()
    Dim X As Integer
    Dim Y As Integer
    For X = 0 To 19
       For Y = X + 1 To 20
          If Collide(shapeInit(X).X + shapeInit(X).xIncrement, shapeInit(X).Y, _
                     shapeInit(Y).X + shapeInit(Y).xIncrement, shapeInit(Y).Y, 750) Then
    
             shapeInit(X).xIncrement = shapeInit(X).xIncrement * -1
             shapeInit(Y).xIncrement = shapeInit(Y).xIncrement * -1
          End If
          If Collide(shapeInit(X).X, shapeInit(X).Y + shapeInit(X).yIncrement, _
                     shapeInit(Y).X, shapeInit(Y).Y + shapeInit(Y).yIncrement, 750) Then
          
             shapeInit(X).yIncrement = shapeInit(X).yIncrement * -1
             shapeInit(Y).yIncrement = shapeInit(Y).yIncrement * -1
          End If
       Next Y
    Next X
    
    For X = 0 To 20
       If (shapeInit(X).X + shapeInit(X).xIncrement + 750) > frmSplash.Width + 650 Or (shapeInit(X).X + shapeInit(X).xIncrement - 750) < 0 Then
          shapeInit(X).xIncrement = shapeInit(X).xIncrement * -1
       End If
      
       If (shapeInit(X).Y + shapeInit(X).yIncrement + 750) > frmSplash.Height + 195 Or (shapeInit(X).Y + shapeInit(X).yIncrement - 750) < 0 Then
          shapeInit(X).yIncrement = shapeInit(X).yIncrement * -1
       End If
       shapeInit(X).X = shapeInit(X).X + shapeInit(X).xIncrement
       shapeInit(X).Y = shapeInit(X).Y + shapeInit(X).yIncrement
       
       Shape1(X).Left = shapeInit(X).X - 750
       Shape1(X).Top = shapeInit(X).Y - 750
    Next X
End Sub
