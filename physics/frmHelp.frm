VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H80000007&
   Caption         =   "Help"
   ClientHeight    =   9075
   ClientLeft      =   4365
   ClientTop       =   1710
   ClientWidth     =   8295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   8295
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   $"frmHelp.frx":0000
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   7935
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tmrState As Boolean

Private Sub cmdOk_Click()
   Me.Hide
   frmMain.tmrMotion = tmrState
   frmMain.Show
End Sub
