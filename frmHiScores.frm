VERSION 5.00
Begin VB.Form frmScores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Best Times"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset Scores"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame fmScores 
      Caption         =   "Fastest Mine Sweepers"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.Label lblHiPlayer 
         Caption         =   "Anonymous"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblHiPlayer 
         Caption         =   "Anonymous"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblHiPlayer 
         Caption         =   "Anonymous"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblTime 
         Caption         =   "999 seconds"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblTime 
         Caption         =   "999 seconds"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblTime 
         Caption         =   "999 seconds"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblLevel 
         Caption         =   "Expert"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblLevel 
         Caption         =   "Intermediate"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblLevel 
         Caption         =   "Beginner"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowScores()
    PrintScores
    Show vbModal
End Sub

Sub PrintScores()
Dim i As Integer
    For i = 0 To 2
        lblTime(i) = Val(GetProfileString(H_TIME & i, "999")) & " seconds"
        lblHiPlayer(i) = GetProfileString(H_NAME & i, "Anonymous")
    Next i
End Sub

Public Sub CheckScores(Level As Levels, Time As Integer)
Dim PlayerName As String
    If Level = Custom Then Exit Sub
    PrintScores
    If Time < GetProfileString(H_TIME & Level, "999") Then
        PlayerName = Left(InputBox("You have the fastest time for the " & lblLevel(Level) & " level. Please insert your name:", "Congratulations", lblHiPlayer(Level)), 15)
        lblTime(Level) = Time & " seconds"
        lblHiPlayer(Level) = PlayerName
        WriteProfileString H_TIME & Level, CStr(Time)
        WriteProfileString H_NAME & Level, PlayerName
        Show vbModal
    End If
End Sub

Private Sub cmdOkay_Click()
    PlaySound SND_CLICK
    Unload Me
End Sub

Private Sub cmdReset_Click()
Dim i As Integer
    PlaySound SND_CLICK
    For i = 0 To 2
        WriteProfileString H_TIME & i, "999"
        WriteProfileString H_NAME & i, "Anonymous"
        lblTime(i) = "999 seconds"
        lblHiPlayer(i) = "Anonymous"
    Next i
End Sub
