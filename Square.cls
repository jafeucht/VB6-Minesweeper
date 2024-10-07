VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Square"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim SControl As Object

Dim Pos As PointAPI
Dim SquarePicture As StdPicture

Dim SurTag As Integer
Dim HTag As Boolean
Dim StatTag As States
Dim Vis As Boolean

Sub Move(Left As Integer, Top As Integer)
    Pos.X = Left
    Pos.Y = Top
    DrawBox
End Sub

Property Get Visible() As Boolean
    Visible = Vis
End Property

Property Let Visible(NewValue As Boolean)
    Vis = NewValue
    DrawBox
End Property

Property Get Left() As Integer
    Left = Pos.X
End Property

Property Let Left(NewValue As Integer)
    Pos.X = NewValue
    DrawBox
End Property

Property Get Top() As Integer
    Top = Pos.Y
End Property

Property Let Top(NewValue As Integer)
    Pos.Y = NewValue
    DrawBox
End Property

Sub DrawBox()
    On Error Resume Next
    If Not Vis Then
        SControl.ForeColor = &H8000000E
        SControl.Line (Pos.X, Pos.Y)-(Pos.X + BOX_SIZE, Pos.Y + BOX_SIZE)
        Exit Sub
    End If
    SControl.PaintPicture SquarePicture, Pos.X, Pos.Y, BOX_SIZE, BOX_SIZE, 0, , BOX_SIZE, BOX_SIZE, vbSrcCopy
End Sub

Property Get SourceControl() As Object
    SourceControl = SControl
End Property

Property Set SourceControl(NewValue As Object)
    Set SControl = NewValue
    DrawBox
End Property

Property Get Picture() As StdPicture
    Set Picture = SquarePicture
End Property

Property Set Picture(PicHandle As StdPicture)
    Set SquarePicture = PicHandle
    DrawBox
End Property

Property Get StateTag() As States
    StateTag = StatTag
End Property

Property Let StateTag(NewValue As States)
    StatTag = NewValue
End Property

Property Get SurroundTag() As Integer
    SurroundTag = SurTag
End Property

Property Let SurroundTag(NewValue As Integer)
    SurTag = NewValue
End Property

Property Get HitTag() As Boolean
    HitTag = HTag
End Property

Property Let HitTag(NewValue As Boolean)
    HTag = NewValue
End Property
