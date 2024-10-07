VERSION 5.00
Begin VB.Form frmField 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Field"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3735
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtMines 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtWidth 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtHeight 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblMaxMines 
      Caption         =   "Max Mines:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblMines 
      Caption         =   "&Mines:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblWidth 
      Caption         =   "&Width:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblHeight 
      Caption         =   "&Height:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Accepted As Boolean
Public varHeight As Integer
Public varWidth As Integer
Public varMineCount As Integer

Const MINE_STR = "Max Mine Quantity: "

Private Sub cmdCancel_Click()
    PlaySound SND_CLICK
    Unload Me
End Sub

Function CheckTextBox(TextRef As TextBox, MinVal As Integer, MaxVal As Integer, Optional FocusText As Boolean = True) As Boolean
Dim BadData As Boolean
    If Not IsNumeric(TextRef) Then
        BadData = True
    ElseIf Val(TextRef) < MinVal Or Val(TextRef) > MaxVal Then
        BadData = True
        If FocusText Then
            TextRef.SelStart = 0
            Beep
            TextRef.SelLength = Len(TextRef)
        End If
    End If
    CheckTextBox = Not BadData
    TextRef.ForeColor = 255 * -BadData
End Function

Private Sub cmdOkay_Click()
    PlaySound SND_CLICK
    If Not CheckTextBox(txtWidth, 8, 55) Then
        MsgBox "Boxes per row must be a numeric value above 8 and below 55."
        Exit Sub
    ElseIf Not CheckTextBox(txtHeight, 8, 35) Then
        MsgBox "Boxes per column must be a numeric value above 8 and below 35."
        Exit Sub
    ElseIf Not CheckTextBox(txtMines, 5, txtWidth * txtHeight - 5) Then
        MsgBox "Number of mines must be a numeric value above 5 and below the maximum mine limit."
        Exit Sub
    End If
    varHeight = txtHeight
    varWidth = txtWidth
    varMineCount = txtMines
    Accepted = True
    Unload Me
End Sub

Private Sub Form_Load()
    txtHeight = BoxesInColumn
    txtWidth = BoxesInRow
    txtMines = MineCount
    Accepted = False
End Sub

Private Sub txtHeight_Change()
    CheckTextBoxes
End Sub

Sub CheckTextBoxes()
    On Error Resume Next
    CheckTextBox txtHeight, 8, 35, False
    CheckTextBox txtWidth, 8, 55, False
    CheckTextBox txtMines, 5, Val(txtHeight) * Val(txtWidth) - 5, False
    lblMaxMines = MINE_STR & Val(txtHeight) * Val(txtWidth) - 5
End Sub

Private Sub txtMines_Change()
    CheckTextBoxes
End Sub

Private Sub txtWidth_Change()
    CheckTextBoxes
End Sub
