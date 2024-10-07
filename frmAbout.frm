VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Minesweeper"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblHLink 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "http://www.geocities.com/Digitronix/"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   600
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "An authentic replica by Jonathan Allen Feucht"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   285
      TabIndex        =   1
      Top             =   1080
      Width           =   3285
   End
   Begin VB.Image imgMine 
      Height          =   480
      Index           =   1
      Left            =   3120
      Top             =   300
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "MINESWEEPER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   750
      TabIndex        =   0
      Top             =   360
      Width           =   2325
   End
   Begin VB.Image imgMine 
      Height          =   480
      Index           =   0
      Left            =   240
      Top             =   300
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Sub cmdClose_Click()
    PlaySound SND_CLICK
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
    For i = 0 To 1
        imgMine(i).Picture = GetResPicture(100, ResIcon)
    Next i
End Sub

Private Sub lblHLink_Click()
    ' This web sight is currently not up, but the following code serves as an example of how to reference a web sight through code
    ShellExecute GetDesktopWindow, vbNullString, "http://www.geocities.com/Digitronix/", vbNullString, vbNullString, vbNormalFocus
End Sub
