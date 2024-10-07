VERSION 5.00
Begin VB.Form frmMine 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minesweeper"
   ClientHeight    =   4170
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4185
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   278
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   279
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctTemp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3720
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pctFaces 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   3720
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pctNumbers 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4140
      Left            =   3480
      ScaleHeight     =   276
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pctSquares 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   3240
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmTimer 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox pctPanel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   1
      Tag             =   "\4"
      Top             =   120
      Width           =   2775
      Begin VB.PictureBox pctFace 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   1080
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   7
         Top             =   60
         Width           =   360
      End
      Begin VB.PictureBox pctTimer 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   1800
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   6
         Tag             =   "/1"
         Top             =   60
         Width           =   585
      End
      Begin VB.PictureBox pctMines 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   60
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   5
         Tag             =   "/1"
         Top             =   60
         Width           =   585
      End
   End
   Begin VB.PictureBox pctField 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H8000000E&
      Height          =   2640
      Left            =   120
      ScaleHeight     =   176
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   190
      TabIndex        =   0
      Tag             =   "\4"
      Top             =   720
      Width           =   2850
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGamePause1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameBeginner 
         Caption         =   "&Beginner"
      End
      Begin VB.Menu mnuGameIntermediate 
         Caption         =   "&Intermediate"
      End
      Begin VB.Menu mnuGameExpert 
         Caption         =   "&Expert"
      End
      Begin VB.Menu mnuGameCustom 
         Caption         =   "&Custom..."
      End
      Begin VB.Menu mnuGamePause2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameMarks 
         Caption         =   "&Marks (?)"
      End
      Begin VB.Menu mnuGameColor 
         Caption         =   "Co&lor"
      End
      Begin VB.Menu mnuGameSound 
         Caption         =   "&Sound"
      End
      Begin VB.Menu mnuGamePause3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameViewHi 
         Caption         =   "&View High Scores"
      End
      Begin VB.Menu mnuGamePause4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu mnuHelpPause 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Box() As New Square

Dim InGame As Boolean
Dim GameOver As Boolean
Dim FaceIndex As Faces
Dim IsSetup As Boolean

Dim RightDown As Boolean
Dim LeftDown As Boolean
Dim BothDown As Boolean
Dim MouseOver As Integer
Dim CurLevel As Levels
Dim Edited As Boolean

Dim CurTime As Integer
Dim CurMineCnt As Integer

Sub SetFace(FaceIdx As Faces, Optional Record As Boolean = True)
Const FACE_SIDE = 24
    If Record Then FaceIndex = FaceIdx
    BitBlt pctFace.hDC, 0, 0, FACE_SIDE, FACE_SIDE, pctFaces.hDC, 0, FaceIdx * FACE_SIDE, vbSrcCopy
End Sub

Sub SetupBlocks()
Dim Pt As PointAPI, CurPt As PointAPI, i As Integer
Dim FieldSize As PointAPI, TmpInt As Integer
    On Error Resume Next
    ReDim Box(1 To BoxesInColumn * BoxesInRow)
    
    FieldSize = GetPoint(BoxesInRow * 16, BoxesInColumn * 16)
    If Not IsSetup Then
        TmpInt = ScaleX(Height, 1, 3) - ScaleHeight
        pctPanel.Move 4, 4, FieldSize.x
        pctFace.Left = pctPanel.Width / 2 - pctFace.Width / 2
        pctTimer.Left = pctPanel.Width - pctTimer.Width - 5
        pctField.Move 4, pctPanel.Top + pctPanel.Height + 10, pctPanel.Width, FieldSize.y
        Move Left, Top, ToTwips(pctPanel.Width + 14), ToTwips(pctField.Top + pctField.Height + TmpInt + 4)
        Move Screen.Width / 2 - Width / 2, Screen.Height / 2 - Height / 2
        
        IsSetup = True
        MakeControls3D Me
    End If
    
    For Pt.y = 1 To BoxesInColumn
        For Pt.x = 1 To BoxesInRow
            i = i + 1
            Box(i).Move CurPt.x, CurPt.y
            
            SetSquare i, 0
            
            Box(i).StateTag = Blank
            Box(i).Visible = True
            Box(i).SurroundTag = 0
            Set Box(i).SourceControl = pctField
            CurPt.x = CurPt.x + BOX_SIZE
        Next Pt.x
        CurPt.x = 0
        CurPt.y = CurPt.y + BOX_SIZE
    Next Pt.y
    
    GameOver = False
    InGame = False
    SetFace Smile
    BothDown = False
    CurMineCnt = MineCount
    ChangeTimer 0
    RefreshDisp pctMines, CurMineCnt
End Sub

Private Function SetSquare(Index As Integer, SquareNum As Integer) As IPictureDisp
Const SIDE_LEN = 16
    BitBlt pctField.hDC, Box(Index).Left, Box(Index).Top, SIDE_LEN, SIDE_LEN, pctSquares.hDC, 0, SIDE_LEN * SquareNum, vbSrcCopy
    Set Box(Index).Picture = pctTemp
End Function

Sub ChangeTimer(TimerVal As Integer)
    CurTime = TimerVal
    If CurTime > 999 Then CurTime = 999
    RefreshDisp pctTimer, CurTime
End Sub

Sub CreateMines()
Dim i As Integer, j As Integer
    Randomize Timer
    For i = 1 To MineCount
        Do
            j = Int(Rnd * (BoxCount)) + 1
        Loop Until Box(j).SurroundTag = 0
        Box(j).SurroundTag = -1
    Next i
    'Debug.Print i
    For i = 1 To BoxCount
        If Box(i).SurroundTag > -1 Then Box(i).SurroundTag = GetSurroundingMines(i)
        'Debug.Print Box(i).SurroundTag
    Next i
    tmTimer.Enabled = True
    InGame = True
End Sub

Function GetSurroundingMines(Index As Integer, Optional CheckFlag As Boolean) As Integer
Dim i As Integer
    
    On Error GoTo SkipHit
    If IsMine(Index - BoxesInRow, CheckFlag) Then i = i + 1
    If IsMine(Index + BoxesInRow, CheckFlag) Then i = i + 1
    If Index Mod BoxesInRow <> 1 Then ' Left edge
        If IsMine(Index - BoxesInRow - 1, CheckFlag) Then i = i + 1
        If IsMine(Index - 1, CheckFlag) Then i = i + 1
        If IsMine(Index + BoxesInRow - 1, CheckFlag) Then i = i + 1
    End If
    If Index Mod BoxesInRow <> 0 Then ' Right edge
        If IsMine(Index - BoxesInRow + 1, CheckFlag) Then i = i + 1
        If IsMine(Index + 1, CheckFlag) Then i = i + 1
        If IsMine(Index + BoxesInRow + 1, CheckFlag) Then i = i + 1
    End If
    GetSurroundingMines = i
    Exit Function
    
SkipHit:
    i = i - 1
    Resume Next
End Function

Function IsMine(Index As Integer, Optional CheckFlag As Boolean) As Boolean
    On Error Resume Next
    If CheckFlag Then
        IsMine = (Box(Index).StateTag = Flag)
    Else
        IsMine = (Box(Index).SurroundTag = -1)
    End If
End Function

Sub ResetHit()
Dim i As Integer
    For i = 1 To BoxCount
        Box(i).HitTag = False
    Next i
End Sub

Sub RevealBoard()
Dim i As Integer
    For i = 1 To BoxCount
        If Box(i).SurroundTag = -1 And Not Box(i).StateTag = Flag Then
            SetSquare i, 5
            Box(i).StateTag = Mine
        ElseIf Box(i).StateTag = Flag And Not Box(i).SurroundTag = -1 Then
            SetSquare i, 4
            Box(i).StateTag = NoMine
        End If
    Next i
End Sub

Sub Uncover(Index As Integer)
    If GameOver Then Exit Sub
    If Index < 1 Or Index > BoxCount Then Exit Sub
    If Box(Index).HitTag Then Exit Sub
    If Box(Index).StateTag = Flag Then Exit Sub
    If Box(Index).StateTag = Number Then Exit Sub
    If Box(Index).StateTag = Blank Then Edited = True
    Box(Index).HitTag = True
    If Not InGame Then
        Box(Index).SurroundTag = 1
        CreateMines
    End If
    If Box(Index).SurroundTag < 0 Then
        SetFace Die
        PlaySound SND_BLAST
        RevealBoard
        
        SetSquare Index, 3
        Box(Index).StateTag = SelMine
        
        
        InGame = False
        GameOver = True
        Exit Sub
    ElseIf Box(Index).SurroundTag > 0 Then
        Box(Index).StateTag = Number
        SetSquare Index, 15 - Box(Index).SurroundTag
    Else
        Box(Index).StateTag = Number
        SetSquare Index, 15
        Uncover Index - BoxesInRow
        Uncover Index + BoxesInRow
        If Index Mod BoxesInRow <> 1 Then ' Left edge
            Uncover Index - BoxesInRow - 1
            Uncover Index - 1
            Uncover Index + BoxesInRow - 1
        End If
        If Index Mod BoxesInRow <> 0 Then ' Right edge
            Uncover Index - BoxesInRow + 1
            Uncover Index + 1
            Uncover Index + BoxesInRow + 1
        End If
    End If
End Sub

Sub RedrawSurroundingBoxes(BoxNumber As Integer, MouseOpp As Integer, Optional Uncover As Integer)
    On Error Resume Next
    RefreshBox BoxNumber - BoxesInRow, MouseOpp, Uncover
    RefreshBox BoxNumber, MouseOpp, Uncover
    RefreshBox BoxNumber + BoxesInRow, MouseOpp, Uncover
    If BoxNumber Mod BoxesInRow <> 1 Then ' Left edge
        RefreshBox BoxNumber - BoxesInRow - 1, MouseOpp, Uncover
        RefreshBox BoxNumber - 1, MouseOpp, Uncover
        RefreshBox BoxNumber + BoxesInRow - 1, MouseOpp, Uncover
    End If
    If BoxNumber Mod BoxesInRow <> 0 Then ' Right edge
        RefreshBox BoxNumber - BoxesInRow + 1, MouseOpp, Uncover
        RefreshBox BoxNumber + 1, MouseOpp, Uncover
        RefreshBox BoxNumber + BoxesInRow + 1, MouseOpp, Uncover
    End If
End Sub

Sub AddFlag(Optional Amount As Integer = 1)
    CurMineCnt = CurMineCnt + Amount
    RefreshDisp pctMines, CurMineCnt
End Sub

Function BoxCount() As Integer
    On Error Resume Next
    BoxCount = UBound(Box)
End Function

Function CheckWin() As Boolean
Dim i As Integer, Cnt As Integer
    If Not InGame Then Exit Function
    For i = 1 To BoxCount
        If Box(i).StateTag = Number Then Cnt = Cnt + 1
    Next i
    If Cnt = BoxesInRow * BoxesInColumn - MineCount Then
        For i = 1 To BoxCount
            If Box(i).SurroundTag = -1 Then Box(i).StateTag = Flag
            RefreshBox i
        Next i
        SetFace Won
        GameOver = True
        InGame = False
        PlaySound SND_WIN
        frmScores.CheckScores CurLevel, CurTime
    End If
End Function

Sub AutoWin()
Dim i As Integer
    For i = 1 To BoxCount
        If Box(i).SurroundTag >= 0 And Box(i).StateTag = Blank Then
            Uncover i
            If GameOver Then Exit Sub
        End If
    Next i
End Sub

Sub RefreshBox(Index As Integer, Optional MouseOpp As Integer, Optional UncoverBox As Integer)
    On Error Resume Next
    If UncoverBox Then
        If GameOver Then Exit Sub
        Uncover (Index)
    End If
    Select Case Box(Index).StateTag
        Case Blank
            SetSquare Index, 15 * MouseOpp
        Case Flag
            SetSquare Index, 1
        Case Question
            SetSquare Index, 2 + 4 * MouseOpp
        Case SelMine
            SetSquare Index, 3
        Case NoMine
            SetSquare Index, 4
        Case Mine
            SetSquare Index, 5
        Case Number
            SetSquare Index, 15 - Box(Index).SurroundTag
    End Select
End Sub

Private Sub Form_Load()
    AppPath = App.Path
    Icon = GetResPicture(100, ResIcon)
    If Not Right(AppPath, 1) = "\" Then AppPath = AppPath & "\"
    GetSettings
    mnuGameMarks.Checked = AllowQMarks
    mnuGameColor.Checked = AllowColor
    mnuGameSound.Checked = AllowSound
    RefreshBoard
    SetFace Smile
    SetupBlocks
    SetLevel BoxesInRow, BoxesInColumn, MineCount
End Sub

Sub RefreshDisp(PictDest As PictureBox, Data As Integer)
Dim DataStr As String * 3, i As Integer, PicNum As Integer, TmpStr As String
Const NUM_HEIGHT = 23
Const NUM_WIDTH = 13
    
    DataStr = Format(Abs(Data), "000")
    If Data < Abs(Data) Then Mid(DataStr, 1, 1) = "-"
    
    For i = 0 To 2
        PicNum = 0
        TmpStr = Mid(DataStr, i + 1, 1)
        If IsNumeric(TmpStr) Then
            PicNum = 11 - TmpStr
        End If
        DoEvents
        BitBlt PictDest.hDC, NUM_WIDTH * i, 0, NUM_WIDTH, NUM_HEIGHT, pctNumbers.hDC, 0, NUM_HEIGHT * PicNum, vbSrcCopy
        PictDest.Refresh
    Next i
End Sub

Function ToTwips(PixelUnits As Single) As Single
    ToTwips = ScaleX(PixelUnits, 3, 1)
End Function

Private Sub Form_Paint()
    MakeControls3D Me
End Sub

Private Sub Form_Resize()
    pctPanel.Cls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteSettings
    End
End Sub

Private Sub mnuGameColor_Click()
    PlaySound SND_CLICK
    AllowColor = Not AllowColor
    mnuGameColor.Checked = AllowColor
    RefreshBoard
End Sub

Sub RefreshBoard()
Dim i As Integer
    pctSquares = GetResPicture(410 + CInt(AllowColor) + 1, ResBitmap)
    pctNumbers = GetResPicture(420 + CInt(AllowColor) + 1, ResBitmap)
    pctFaces = GetResPicture(430 + CInt(AllowColor) + 1, ResBitmap)
    RefreshDisp pctTimer, CurTime
    RefreshDisp pctMines, CurMineCnt
    SetFace FaceIndex
    For i = 1 To BoxesInRow * BoxesInColumn
        RefreshBox i
    Next i
End Sub

Private Sub mnuGameNew_Click()
    PlaySound SND_CLICK
    pctFace_Click
End Sub

Private Sub mnuGameSound_Click()
    PlaySound SND_CLICK
    AllowSound = Not AllowSound
    mnuGameSound.Checked = AllowSound
End Sub

Private Sub mnuHelpAbout_Click()
    PlaySound SND_CLICK
    frmAbout.Show vbModal
End Sub

Private Sub pctFace_Click()
    PlaySound SND_CLICK
    SetupBlocks
End Sub

Private Sub pctFace_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetFace MouseDwn, False
End Sub

Private Sub pctFace_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetFace FaceIndex
End Sub

Private Sub pctFace_Paint()
    SetFace FaceIndex
End Sub

Private Sub mnuGameBeginner_Click()
    PlaySound SND_CLICK
    SetLevel 8, 8, 10
End Sub

Private Sub mnuGameCustom_Click()
    PlaySound SND_CLICK
    frmField.Show vbModal
    If frmField.Accepted Then
        SetLevel frmField.varWidth, frmField.varHeight, frmField.varMineCount
    End If
End Sub

Private Sub mnuGameExit_Click()
    PlaySound SND_CLICK
    Unload Me
End Sub

Private Sub mnuGameExpert_Click()
    PlaySound SND_CLICK
    SetLevel 30, 16, 99
End Sub

Private Sub mnuGameIntermediate_Click()
    PlaySound SND_CLICK
    SetLevel 16, 16, 40
End Sub

Sub AdjustLevelMenus(BoxRow, BoxCol, Mines)
    mnuGameBeginner.Checked = False
    mnuGameIntermediate.Checked = False
    mnuGameExpert.Checked = False
    mnuGameCustom.Checked = False
    If BoxRow = 8 And BoxCol = 8 And Mines = 10 Then
        CurLevel = Beginner
        mnuGameBeginner.Checked = True
    ElseIf BoxRow = 16 And BoxCol = 16 And Mines = 40 Then
        CurLevel = Intermediate
        mnuGameIntermediate.Checked = True
    ElseIf BoxRow = 30 And BoxCol = 16 And Mines = 99 Then
        CurLevel = Expert
        mnuGameExpert.Checked = True
    Else
        CurLevel = Custom
        mnuGameCustom.Checked = True
    End If
End Sub

Sub SetLevel(BoxRow, BoxCol, Mines)
    InGame = False
    BoxesInRow = BoxRow
    BoxesInColumn = BoxCol
    MineCount = Mines
    IsSetup = False
    SetupBlocks
    AdjustLevelMenus BoxRow, BoxCol, Mines
End Sub

Private Function GetBlockNum(Position As PointAPI) As Integer
Dim PosX As Integer, PosY As Integer
    PosX = Position.x \ BOX_SIZE Mod BoxesInRow
    PosY = Position.y \ BOX_SIZE Mod BoxesInColumn
    GetBlockNum = PosY * BoxesInRow + PosX + 1
End Function

Private Sub mnuGameMarks_Click()
    PlaySound SND_CLICK
    AllowQMarks = Not AllowQMarks
    mnuGameMarks.Checked = AllowQMarks
End Sub

Private Sub mnuGameViewHi_Click()
    PlaySound SND_CLICK
    frmScores.ShowScores
End Sub

Private Sub mnuHelpContents_Click()
    ShellExecute hwnd, "Open", AppPath & "WINMINE.CHM", 0, AppPath, 1
    PlaySound SND_CLICK
End Sub

Private Sub pctField_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Index As Integer
    Index = GetBlockNum(GetPoint(CInt(x), CInt(y)))
    Select Case Button
        Case 1
            LeftDown = True
        Case 2
            RightDown = True
        Case Else
            Exit Sub    ' Only supports two mouse buttons
    End Select
    BothDown = RightDown And LeftDown
    If OutOfBox(x, y) Or GameOver Then Exit Sub
    RedrawSurroundingBoxes Index, Abs(BothDown)
    If RightDown And Not LeftDown And Not BothDown Then
        Select Case Box(Index).StateTag
            Case Blank
                Box(Index).StateTag = Flag
                AddFlag -1
            Case Flag
                AddFlag
                If AllowQMarks Then
                    Box(Index).StateTag = Question
                Else
                    Box(Index).StateTag = Blank
                End If
            Case Question
                Box(Index).StateTag = Blank
        End Select
        'pctMines =
        RefreshBox Index, 0
    ElseIf LeftDown And Not RightDown And Not BothDown Then
        SetFace Suspense
        RefreshBox Index, 1
    End If
End Sub

Function OutOfBox(x As Single, y As Single) As Boolean
    OutOfBox = x >= pctField.Width Or y >= pctField.Height Or x < 0 Or y < 0
End Function

Private Sub pctField_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Index As Integer
    Index = GetBlockNum(GetPoint(CInt(x), CInt(y)))
    If BothDown And Not (LeftDown And RightDown) Then Exit Sub
    If Not LeftDown Then Exit Sub
    If Not (LeftDown Or RightDown) Then Exit Sub
    RedrawSurroundingBoxes MouseOver, 0
    If OutOfBox(x, y) Or GameOver Then Exit Sub
    If LeftDown And RightDown Then
        'Debug.Print "Hi"
        RedrawSurroundingBoxes Index, 1
    End If
    RefreshBox Index, 1
    MouseOver = GetBlockNum(GetPoint(CInt(x), CInt(y)))
End Sub

Private Sub pctField_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Index As Integer, i As Integer, IsPressed As Boolean
    Index = GetBlockNum(GetPoint(CInt(x), CInt(y)))
    Edited = False
    Select Case Button
        Case 1
            LeftDown = False
        Case 2
            RightDown = False
        Case Else
            Exit Sub    ' Only supports two mouse buttons
    End Select
    IsPressed = Box(Index).StateTag > Blank
    If Not RightDown And Not LeftDown Then
        BothDown = False
    End If
    If OutOfBox(x, y) Or GameOver Then
        If Not GameOver Then SetFace Smile
        Exit Sub
    End If
    If BothDown And RightDown Xor LeftDown Then
        RedrawSurroundingBoxes Index, 0, Box(Index).StateTag = Number And GetSurroundingMines(Index, True) >= Box(Index).SurroundTag
    ElseIf Button = 1 And Not BothDown Then
        Uncover Index
    End If
    CheckWin
    If Not GameOver Then
        If Edited Then MakeStep
        SetFace Smile
    End If
End Sub

Private Sub pctPanel_Paint()
    MakeControls3D Me
End Sub

Private Sub tmTimer_Timer()
    If Not InGame Then
        tmTimer.Enabled = False
        Exit Sub
    End If
    ChangeTimer CurTime + 1
End Sub
