VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   689
   WindowState     =   2  'Maximized
   Begin VB.Frame frmAttributes 
      Caption         =   "Planet Attributes"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   2775
      Begin VB.CommandButton cmdResumPause 
         Caption         =   "Pause/Resume"
         Enabled         =   0   'False
         Height          =   495
         Left            =   480
         TabIndex        =   29
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label lblDoom 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   34
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Time until the Death Star is full charged:"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   33
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lblDist 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   31
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Distance form the sun:"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   30
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "km/s"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   22
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblAcel 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   27
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblAcelAng 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblSpAng 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   21
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "km/s"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblSpeed 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Velocity of the planet:"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label16 
         Caption         =   "Acceleration of the planet:"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Frame frmSettings 
      Caption         =   "Settings"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   12600
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox txtAngle 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "0"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtSpeed 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "29.78"
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox txtDistance 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Text            =   "149.6"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtMass 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Text            =   "1.98892"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   -120
         X2              =   2280
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2400
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Initial velocity of the planet"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Distance from planet to star"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   15
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "km/s"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   14
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   13
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   12
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "x 10"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   11
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "g"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   10
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   9
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Mass of the star"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "x 10"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Timer tmrMotion 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   4800
   End
   Begin VB.Menu mnuFIle 
      Caption         =   "File"
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuPresets 
      Caption         =   "Presets"
      Begin VB.Menu mnuMercury 
         Caption         =   "Mercury"
      End
      Begin VB.Menu mnuVenus 
         Caption         =   "Venus"
      End
      Begin VB.Menu mnuEarth 
         Caption         =   "Earth"
      End
      Begin VB.Menu mnuMars 
         Caption         =   "Mars"
      End
      Begin VB.Menu mnuCir 
         Caption         =   "Perfect circular motion"
      End
      Begin VB.Menu mnuC1 
         Caption         =   "Comet1"
      End
      Begin VB.Menu mnuC2 
         Caption         =   "Comet 2"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuKeplar 
         Caption         =   "Show kepler's law"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Free Play"
      End
      Begin VB.Menu mnuTime 
         Caption         =   "Time interval"
         Begin VB.Menu mnuT1 
            Caption         =   "1 ms"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuT100 
            Caption         =   "100 ms"
         End
         Begin VB.Menu mnuT250 
            Caption         =   "250 ms"
         End
      End
   End
   Begin VB.Menu mnuHlp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'all the declerations Sau.J
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Dim pen As Long
Dim cnt As Integer 'used for kepler's law Sau.J
Dim kep As Boolean 'should we show kepler's law? Sau.J
Dim x1, y1 As Long 'used for draw orbit of the planet. Sau.J
Dim x2, y2 As Long 'used to draw the deathstar's orbit. Sau.J
Dim doom As Integer 'time till the death star fires Sau.J
Private Type BitMap 'stores the planet images Sau.J
    Type As Long
    Width As Long
    Height As Long
    WidthBytes As Long
    Planes As Integer
    BitsPixel As Integer
    Bits As Long
    hDC As Long
End Type
'used for double buffering Sau.J
Dim backBuff As Long
'universal gravitaion constance Sau.J
Const G = 6.67 * 10 ^ -11
'sun mass; not constant since it varies Sau.J
Dim sMass As Double
'self explainatory Sau.J
Private Type planet
   mass As Long
   distance As Double
   velX As Double
   velY As Double
   angle As Double
   x As Double
   Y As Double
   accel As Double
   pic As BitMap
End Type

Dim freePlay As Boolean 'if the user just wants to goof around Sau.J
Dim curOrbit As Byte
Dim orbits(1 To 6) As planet  'all possible orbits that the deathstar can have Sau.J

'the planet that the user plays around with Sau.J
Dim myPlanet As planet
Dim deathStar As planet
Dim sun As BitMap 'sun's picture Sau.J
'a very helpful function Sau.J
Private Function randInt(ByVal x As Integer, ByVal Y As Integer)
   If x > Y Then
      Error (3)
   End If
   randInt = Int(Rnd * (Y - x + 1)) + x
End Function
Function distance(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
   'finds the distace between two ponts Sau.J
   distance = ((x1 - x2) ^ 2 + (y1 - y2) ^ 2) ^ 0.5 'square root is something to the 1/2 Sau.J
End Function

Function angle(px As Double, py As Double, tx As Double, ty As Double) As Double
 'finds the angle between two point in degrees Sau.J
 'uses the tangent inverse function Sau.J
 'very very messy Sau.J
 Dim adjecent As Double
 Dim opposite As Double
 adjecent = tx - px
 opposite = ty - py
   
 If adjecent = 0 And ty > py Then
    angle = 90
    Exit Function
 ElseIf adjecent = 0 And ty <= py Then
    angle = 270
    Exit Function
 End If
 
 If opposite = 0 And tx > px Then
    angle = 0
    Exit Function
 End If
        
 If opposite = 0 And tx < px Then
    angle = 180
    Exit Function
 End If
 'I got confused, so I brute-forced this part Sau.J
 If ty < py Then
    If Atn(opposite / adjecent) * 180 / 3.14159265 < 0 Then
        angle = (Atn(opposite / adjecent) * 180 / 3.14159265) 'to convert radian in to degrees you must multiply by 180/PI Sau.J
        Exit Function
    Else
        angle = (Atn(opposite / adjecent) * 180 / 3.14159265) - 180
        Exit Function
    End If
    angle = 180 + Atn(opposite / adjecent) * 180 / 3.14159265
    Exit Function
 Else
    If Atn(opposite / adjecent) < 0 Then
       angle = 180 + Atn(opposite / adjecent) * 180 / 3.14159265
       Exit Function
    Else
       angle = Atn(opposite / adjecent) * 180 / 3.14159265
       Exit Function
    End If
 End If
End Function
'uses the formular for equation Sau.J
Function acceleration(r As Double) As Double
   acceleration = (G * sMass) / (r * r)
End Function
'x xomponent of the acceleration v
Function pullX(acel As Double, ang As Double) As Double
   pullX = Cos((ang) * 3.14159265 / 180) * acel
End Function
'y component of the acceleration Sau.J
Function pullY(acel As Double, ang As Double) As Double
   pullY = Sin((ang) * 3.14159265 / 180) * -acel
End Function
'procedure to draw pictures Sau.J
Private Sub picDraw(PicInfo As BitMap, x As Long, Y As Long)
   BitBlt backBuff, x, Y, PicInfo.Width, PicInfo.Height, PicInfo.hDC, 0, 0, vbSrcCopy
End Sub
'update the screen Sau.J
Private Sub viewUpdate()
   BitBlt Me.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, backBuff, 0, 0, vbSrcCopy
End Sub
'take the parametres and runs the simulation Sau.J
Private Sub cmdStart_Click()
   sMass = txtMass.Text * 10 ^ 30
   myPlanet.angle = txtAngle.Text
   myPlanet.distance = txtDistance.Text * 10 ^ 6
   myPlanet.x = Me.ScaleWidth / 2 - myPlanet.pic.Width
   myPlanet.Y = Me.ScaleHeight / 2 + txtDistance.Text - myPlanet.pic.Height
   myPlanet.accel = acceleration(myPlanet.distance)
   myPlanet.velX = pullX(txtSpeed.Text * 10 ^ 3, myPlanet.angle)
   myPlanet.velY = pullY(txtSpeed.Text * 10 ^ 3, myPlanet.angle)
   tmrMotion.Enabled = True
   MoveToEx backBuff, myPlanet.x, myPlanet.Y, 0
   x1 = myPlanet.x
   y1 = myPlanet.Y
   BitBlt backBuff, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, 0, vbBlack
   cmdResumPause.Enabled = True
   
   deathStar.angle = orbits(curOrbit).angle
   deathStar.distance = orbits(curOrbit).distance
   deathStar.x = orbits(curOrbit).x
   deathStar.Y = orbits(curOrbit).Y
   deathStar.accel = orbits(curOrbit).accel
   deathStar.velX = orbits(curOrbit).velX
   deathStar.velY = orbits(curOrbit).velY
   doom = 1000
End Sub
'pause/resume Sau.J
Private Sub cmdResumPause_Click()
   tmrMotion.Enabled = Not tmrMotion.Enabled
End Sub
'move the side planels around Sau.J
Private Sub Form_DragDrop(Source As Control, x As Single, Y As Single)
   Source.Left = x
   Source.Top = Y
End Sub

Private Sub Form_Load()
    Randomize
   backBuff = CreateCompatibleDC(GetDC(0))
   sMass = 1.98892 * 10 ^ 30
   myPlanet.pic = PicLoad(App.Path + "\p.bmp")
   sun = PicLoad(App.Path + "\s.bmp")
   pen = CreatePen(0, 1, vbYellow)
   DeleteObject SelectObject(backBuff, pen)
   deathStar.pic = PicLoad("b.bmp")
   freePlay = False

   'load up all the orbits Sau.J
   orbits(1).angle = 0
   orbits(1).distance = 58 * 10 ^ 6
   orbits(1).x = 1024 / 2 - myPlanet.pic.Width
   orbits(1).Y = 722 / 2 + 58 - myPlanet.pic.Height
   orbits(1).accel = acceleration(orbits(1).distance)
   orbits(1).velX = pullX(49 * 10 ^ 3, orbits(1).angle)
   orbits(1).velY = pullY(49 * 10 ^ 3, orbits(1).angle)
   
   orbits(2).angle = 0
   orbits(2).distance = 107.5 * 10 ^ 6
   orbits(2).x = 1024 / 2 - deathStar.pic.Width
   orbits(2).Y = 722 / 2 + 107.5 - deathStar.pic.Height
   orbits(2).accel = acceleration(orbits(2).distance)
   orbits(2).velX = pullX(34.89 * 10 ^ 3, orbits(2).angle)
   orbits(2).velY = pullY(34.89 * 10 ^ 3, orbits(2).angle)
   
   orbits(3).angle = 0
   orbits(3).distance = 228 * 10 ^ 6
   orbits(3).x = 1024 / 2 - deathStar.pic.Width
   orbits(3).Y = 722 / 2 + 228 - deathStar.pic.Height
   orbits(3).accel = acceleration(orbits(3).distance)
   orbits(3).velX = pullX(24 * 10 ^ 3, orbits(3).angle)
   orbits(3).velY = pullY(24 * 10 ^ 3, orbits(3).angle)

   orbits(4).angle = 0
   orbits(4).distance = 250 * 10 ^ 6
   orbits(4).x = 1024 / 2 - deathStar.pic.Width
   orbits(4).Y = 722 / 2 + 250 - deathStar.pic.Height
   orbits(4).accel = acceleration(orbits(3).distance)
   orbits(4).velX = pullX(11.5524889 * 10 ^ 3, orbits(4).angle)
   orbits(4).velY = pullY(11.5524889 * 10 ^ 3, orbits(4).angle)
   
   orbits(5).angle = 0
   orbits(5).distance = 269 * 10 ^ 6
   orbits(5).x = 1024 / 2 - deathStar.pic.Width
   orbits(5).Y = 722 / 2 + 269 - deathStar.pic.Height
   orbits(5).accel = acceleration(orbits(5).distance)
   orbits(5).velX = pullX(15 * 10 ^ 3, orbits(5).angle)
   orbits(5).velY = pullY(15 * 10 ^ 3, orbits(5).angle)
   
   orbits(6).angle = 135
   orbits(6).distance = 269 * 10 ^ 6
   orbits(6).x = 1024 / 2 - deathStar.pic.Width
   orbits(6).Y = 722 / 2 + 269 - deathStar.pic.Height
   orbits(6).accel = acceleration(orbits(6).distance)
   orbits(6).velX = pullX(15 * 10 ^ 3, orbits(6).angle)
   orbits(6).velY = pullY(15 * 10 ^ 3, orbits(6).angle)
   'select the cutten orbit Sau.J
   curOrbit = randInt(1, 6)
   If curOrbit = 6 Then
      txtMass.Text = 0.59
   End If
End Sub
'loads a picture Sau.J
Private Function PicLoad(ByVal FileName As String) As BitMap
   Dim PicInfo As BitMap
   
   PicInfo.hDC = CreateCompatibleDC(GetDC(0))
   SelectObject PicInfo.hDC, LoadPicture(FileName)
   GetObject LoadPicture(FileName), Len(PicInfo), PicInfo
   
   PicLoad = PicInfo
   Exit Function
End Function

Private Sub Form_Resize()
    'Create bitmap of proper size for the background
    Dim backBitmap As Long
    backBitmap = CreateCompatibleBitmap(GetDC(0), Me.ScaleWidth, Me.ScaleHeight)
    SelectObject backBuff, backBitmap
    DeleteObject backBitmap
End Sub
'free the memory, don't want a stack overflow! Sau.J
Private Sub Form_Unload(Cancel As Integer)
    DeleteDC backBuff
    DeleteDC myPlanet.pic.hDC
    DeleteDC sun.hDC
    DeleteDC deathStar.pic.hDC
End Sub

Private Sub mnuHelp_Click()
   frmHelp.tmrState = tmrMotion.Enabled
   Me.Hide
   frmHelp.Show
End Sub

Private Sub mnuPlay_Click()
   mnuPlay.Checked = Not mnuPlay.Checked
   freePlay = Not freePlay
   txtMass.Enabled = Not txtMass.Enabled
   txtMass.Text = 1.98892
   curOrbit = randInt(1, 6)
   If curOrbit = 6 Then
      txtMass.Text = 0.59
   End If
End Sub

Private Sub mnuQuit_Click()
   Dim x As Integer
   x = MsgBox("are you sure you want to quit?", vbYesNo)
   If x = vbYes Then
      End
   End If
End Sub

'changes the time intervals Sau.J
Private Sub mnuT1_Click()
   tmrMotion.Interval = 1
   mnuT1.Checked = True
   mnuT100.Checked = False
   mnuT250.Checked = False
End Sub

Private Sub mnuT100_Click()
   tmrMotion.Interval = 100
   mnuT1.Checked = False
   mnuT100.Checked = True
   mnuT250.Checked = False
End Sub

Private Sub mnuT250_Click()
   tmrMotion.Interval = 250
   mnuT1.Checked = False
   mnuT100.Checked = False
   mnuT250.Checked = True
End Sub

Private Sub mnuC1_Click()
   txtMass.Text = 1.98892
   txtDistance.Text = 269
   txtSpeed.Text = 15
   txtAngle.Text = 0
   
End Sub

Private Sub mnuC2_Click()
   txtMass.Text = 0.59
   txtDistance.Text = 269
   txtSpeed.Text = 15
   txtAngle.Text = 135
End Sub

Private Sub mnuCir_Click()
   txtMass.Text = 0.5
   txtDistance.Text = 250
   txtSpeed.Text = 11.5524889
   txtAngle.Text = 0
End Sub

Private Sub mnuEarth_Click()
   txtMass.Text = 1.98892
   txtDistance.Text = 149.6
   txtSpeed.Text = 29.78
   txtAngle.Text = 0
End Sub

Private Sub mnuKeplar_Click()
   mnuKeplar.Checked = Not mnuKeplar.Checked
End Sub

Private Sub Timer1_Timer()
   kep = Not kep
End Sub
'all the planets and comets Sau.J
Private Sub mnuMars_Click()
   txtMass.Text = 1.98892
   txtDistance.Text = 228
   txtSpeed.Text = 24
   txtAngle.Text = 0
End Sub

Private Sub mnuMercury_Click()
   txtMass.Text = 1.98892
   txtDistance.Text = 58
   txtSpeed.Text = 49
   txtAngle.Text = 0
End Sub

Private Sub mnuVenus_Click()
   txtMass.Text = 1.98892
   txtDistance.Text = 107.5
   txtSpeed.Text = 34.89
   txtAngle.Text = 0
End Sub

Public Sub tmrMotion_Timer()
   'remembers how to draw the path of the planet Sau.J
   x1 = myPlanet.x
   y1 = myPlanet.Y
   x2 = deathStar.x
   y2 = deathStar.Y
   
   'mover the planet Sau.J
   myPlanet.x = myPlanet.x + myPlanet.velX / 10 ^ 3
   myPlanet.Y = myPlanet.Y + myPlanet.velY / 10 ^ 3
   'gets the distance from the sun to the planet Sau.J
   myPlanet.distance = distance(myPlanet.x, myPlanet.Y, Me.ScaleWidth / 2, Me.ScaleHeight / 2) * 10 ^ 6
      lblDist.Caption = myPlanet.distance
   'accelerates the planet towards the sun Sau.J
   myPlanet.accel = acceleration(myPlanet.distance)
      'direction of the planet Sau.J
   myPlanet.angle = 180 - angle(Me.ScaleWidth / 2, Me.ScaleHeight / 2, myPlanet.x, myPlanet.Y)
   
   'changes the velocity of the planet Sau.J
   myPlanet.velX = myPlanet.velX + pullX(myPlanet.accel, myPlanet.angle)
   myPlanet.velY = myPlanet.velY + pullY(myPlanet.accel, myPlanet.angle)
      'whre the drawing happens Sau.J
   Call picDraw(sun, Me.ScaleWidth / 2 - sun.Width / 2, Me.ScaleHeight / 2 - sun.Height / 2)
   If cnt >= 2 Then
      kep = True
      cnt = 0
   End If
   If mnuKeplar.Checked And kep Then
      pen = CreatePen(0, 1, vbBlue)
      DeleteObject SelectObject(backBuff, pen)
      kep = False
      MoveToEx backBuff, Me.ScaleWidth / 2, Me.ScaleHeight / 2, 0
      LineTo backBuff, myPlanet.x, myPlanet.Y
      
      pen = CreatePen(0, 1, vbYellow)
      DeleteObject SelectObject(backBuff, pen)
      MoveToEx backBuff, x1, y1, 0
   End If
   cnt = cnt + 1
   MoveToEx backBuff, x1, y1, 0
   LineTo backBuff, myPlanet.x, myPlanet.Y
   If angle(myPlanet.velX, myPlanet.velY, 0, 0) < 0 Then
      lblSpAng.Caption = angle(myPlanet.velX, myPlanet.velY, 0, 0) + 180
   Else
      lblSpAng.Caption = angle(myPlanet.velX, myPlanet.velY, 0, 0)
   End If
   'just does the same thig but for the deathstar Sau.J
   If Not freePlay Then
      deathStar.x = deathStar.x + deathStar.velX / 10 ^ 3
      deathStar.Y = deathStar.Y + deathStar.velY / 10 ^ 3
      deathStar.distance = distance(deathStar.x, deathStar.Y, Me.ScaleWidth / 2, Me.ScaleHeight / 2) * 10 ^ 6
      deathStar.angle = 180 - angle(Me.ScaleWidth / 2, Me.ScaleHeight / 2, deathStar.x, deathStar.Y)
      deathStar.velX = deathStar.velX + pullX(deathStar.accel, deathStar.angle)
      deathStar.velY = deathStar.velY + pullY(deathStar.accel, deathStar.angle)
      deathStar.accel = acceleration(deathStar.distance)
      If distance(deathStar.x, deathStar.Y, myPlanet.x, myPlanet.Y) < myPlanet.pic.Width Then
        cmdResumPause.Enabled = False
        tmrMotion.Enabled = False
        MsgBox "Yea you win!!"
      End If
      curOrbit = randInt(1, 6)
      If curOrbit = 6 Then
         txtMass.Text = 0.59
      End If
      doom = doom - 1
      lblDoom = doom
      'if the deathstar is ready Sau.J
      If doom <= 0 Then
         cmdResumPause.Enabled = False
         tmrMotion.Enabled = False
         MsgBox "The Death Star is now fully charged, and has fired. All your base are belong to us!!!"
         pen = CreatePen(0, 5, vbGreen)
         DeleteObject SelectObject(backBuff, pen)
      
         MoveToEx backBuff, myPlanet.x, myPlanet.Y, 0
         LineTo backBuff, deathStar.x, deathStar.Y
      
        pen = CreatePen(0, 1, vbYellow)
        DeleteObject SelectObject(backBuff, pen)
      End If
   End If
   'draw the deathstar's orbit Sau.J
   If Not freePlay Then
      pen = CreatePen(0, 1, vbRed)
      DeleteObject SelectObject(backBuff, pen)
      
      MoveToEx backBuff, x2, y2, 0
      LineTo backBuff, deathStar.x, deathStar.Y
      
      pen = CreatePen(0, 1, vbYellow)
      DeleteObject SelectObject(backBuff, pen)
      MoveToEx backBuff, x1, y1, 0
   End If
   lblAcelAng.Caption = myPlanet.angle
   lblAcel.Caption = myPlanet.accel
   'the planet isn't part of the backbuffer Sau.J
   Call viewUpdate
   BitBlt hDC, myPlanet.x - myPlanet.pic.Width / 2, myPlanet.Y - myPlanet.pic.Height / 2, myPlanet.pic.Width, myPlanet.pic.Height, myPlanet.pic.hDC, 0, 0, vbSrcCopy
   If Not freePlay Then
      BitBlt hDC, deathStar.x - deathStar.pic.Width / 2, deathStar.Y - deathStar.pic.Height / 2, deathStar.pic.Width, deathStar.pic.Height, deathStar.pic.hDC, 0, 0, vbSrcCopy
   End If
End Sub

