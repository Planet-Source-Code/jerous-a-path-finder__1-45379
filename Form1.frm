VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Bomberkid's Path Finder"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   FillColor       =   &H00C00000&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Search Paths"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   2730
      TabIndex        =   2
      Top             =   -60
      Width           =   1800
      Begin VB.CommandButton Command4 
         Caption         =   "Clear Waypoints"
         Height          =   330
         Left            =   75
         TabIndex        =   9
         Top             =   1965
         Width           =   1485
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Find way to ..."
         Height          =   495
         Left            =   75
         TabIndex        =   11
         Top             =   165
         Width           =   1485
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Set Waypoints"
         Height          =   330
         Left            =   75
         TabIndex        =   10
         Top             =   1650
         Width           =   1485
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "v"
         Height          =   405
         Left            =   1260
         TabIndex        =   8
         Top             =   2835
         Width           =   315
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "^"
         Height          =   390
         Left            =   1260
         TabIndex        =   7
         Top             =   2340
         Width           =   315
      End
      Begin VB.ListBox List2 
         Height          =   1185
         Left            =   135
         TabIndex        =   6
         Top             =   2325
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "In order"
         Height          =   300
         Left            =   135
         TabIndex        =   5
         Top             =   3570
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   960
         Left            =   105
         TabIndex        =   3
         Top             =   645
         Width           =   1440
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   75
         TabIndex        =   13
         Top             =   5130
         Width           =   1680
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   45
         TabIndex        =   12
         Top             =   3870
         Width           =   1710
      End
   End
   Begin VB.VScrollBar scrY 
      Height          =   5295
      LargeChange     =   10
      Left            =   6840
      Max             =   50
      Min             =   5
      TabIndex        =   1
      Top             =   75
      Value           =   10
      Width           =   270
   End
   Begin VB.HScrollBar scrX 
      Height          =   255
      LargeChange     =   10
      Left            =   15
      Max             =   50
      Min             =   5
      TabIndex        =   0
      Top             =   5460
      Value           =   10
      Width           =   6750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim V() As Boolean
Const F = 400
Dim XP%, YP% 'positie
Dim PrevX%, PrevY%

Dim FindWay As Boolean, SetWayPoints As Boolean
Dim WayPoints As Collection


Private Sub MaakLijstWP()
Dim X%
If List2.ListCount = WayPoints.Count Then Exit Sub
List2.Clear
For X = 1 To WayPoints.Count
  List2.AddItem Trim(Str(WayPoints(X).X)) + "," + Trim(Str(WayPoints(X).Y))
Next
End Sub

Private Sub Check1_Click()
Teken
End Sub

Private Sub cmdDown_Click()
Dim T As Coordinate, Ind%
Ind = List2.ListIndex
If Ind < List2.ListCount - 1 And Ind >= 0 Then
  Set T = WayPoints(Ind + 1)
  WayPoints.Remove Ind + 1
  WayPoints.Add T, , , Ind + 1
  List2.Clear
  MaakLijstWP
  List2.ListIndex = Ind + 1
  Teken
End If
End Sub

Private Sub cmdUp_Click()
Dim T As Coordinate, Ind%
Ind = List2.ListIndex
If Ind > 0 Then
  Set T = WayPoints(Ind + 1)
  WayPoints.Remove Ind + 1
  WayPoints.Add T, , Ind
  List2.Clear
  MaakLijstWP
  List2.ListIndex = Ind - 1
  Teken
End If
End Sub

Private Sub Command1_Click()
Dim Tmr As Long, Co As Coordinate
ClearField

Tmr = GetTickCount
Set Co = CreateCoordinates(XP, YP)
CreatePaths Co
Label2 = "Path Finding=" + Str(GetTickCount - Tmr) + vbCrLf

Tmr = GetTickCount
Teken
Label2 = Label2 + "Draw=" + Str(GetTickCount - Tmr)
End Sub
Sub ClearField()
Dim X%, Y%
For X = 0 To UBound(Field, 1)
  For Y = 0 To UBound(Field, 2)
    If Field(X, Y).Block <> 1 Then
      Field(X, Y).Block = 0
    End If
  Next
Next
End Sub

Private Sub Command2_Click()
List1.Clear
Command2.FontBold = Not (Command2.FontBold)
If Command2.FontBold Then
  Command1.Enabled = False
'  Command3.Enabled = True
Else
'  Command3.Enabled = False
  If Command3.FontBold Then Command2.FontBold = True: Command3.FontBold = False
  SetWayPoints = False
End If
FindWay = (Command2.FontBold)
'Exit Sub
'MsgBox T
End Sub

Private Sub Command3_Click()
Command3.FontBold = Not (Command3.FontBold)
SetWayPoints = Command3.FontBold
End Sub

Private Sub Command4_Click()
Set WayPoints = New Collection
List2.Clear
Teken
End Sub

Private Sub Form_Click()
'Command1_Click
End Sub

Private Sub Form_Load()
ReDim Field(10, 10)
'Dim D$, X%, Y%, Data, Tel%
'D = GetSetting("BomberKid", "Pad Finder", "Field")
'If D <> "" Then
'  Data = Split(D, ";")
'  ReDim Field(Mid(Data(0), 1, 2), Mid(Data(0), 4, 2))
'  XP = Mid(Data(1), 1, InStr(Data(1), ",") - 1)
'  YP = Mid(Data(1), InStr(Data(1), ",") + 1)
''  ToX = Mid(Data(2), 1, InStr(Data(2), ",") - 1)
''  ToY = Mid(Data(2), InStr(Data(2), ",") + 1)
'  For X = 1 To Len(Data(2))
'    Field(Tel, Y).Block = Mid(Data(2), X, 1)
'    Y = Y + 1
'    If Y > UBound(Field, 2) Then
'      Tel = Tel + 1
'      Y = 0
'    End If
'  Next
'End if
Set WayPoints = New Collection
SetPoints
'Teken
Command1_Click
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = Int(X / F)
Y = Int(Y / F)
If FindWay Or SetWayPoints Then Exit Sub
If X > UBound(Field, 1) Or Y > UBound(Field, 2) Then Exit Sub
If Button Then
  Form_MouseMove Button, Shift, X * F, Y * F
  Exit Sub
End If
Teken
End Sub
Sub TekenPad()
Dim X%, Pos
Teken
For X = 0 To List1.ListCount - 1
  Pos = Split(List1.List(X), ",")
  Line (Pos(0) * F, Pos(1) * F)-((Pos(0) + 1) * F, (Pos(1) + 1) * F), vbYellow, BF
Next
Teken -1, -1
End Sub
Sub Teken(Optional XPos, Optional YPos)
Dim X%, Y%, T$, P1%, P2%, D$(), Draw As Boolean, Ind%
Dim Max%, TestCo As Coordinate
For X = 0 To UBound(Field, 1)
  For Y = 0 To UBound(Field, 2)
    If Field(X, Y).Steps > Max Then Max = Field(X, Y).Steps
  Next
Next
FontSize = 11
If Max = 0 Then Max = 1
For X = 0 To UBound(Field, 1)
  For Y = 0 To UBound(Field, 2)
    Draw = False
    If (IsMissing(XPos) Or IsMissing(YPos)) Then
      Draw = True
    Else
      If (X = XPos And Y = YPos) Then
        Draw = True
      End If
    End If
    If Draw Then
      T = ""
      If Field(X, Y).Block >= 2 And Field(X, Y).Block <= 5 Then
        ForeColor = vbBlue
        If Field(X, Y).Block = e_Up Then T = "^"
        If Field(X, Y).Block = e_Right Then T = ">"
        If Field(X, Y).Block = e_Down Then T = "v"
        If Field(X, Y).Block = e_Left Then T = "<"
      End If
      If Field(X, Y).Block = 1 Then
        Line (X * F, Y * F)-((X + 1) * F, (Y + 1) * F), vbBlack, BF
      Else
        Line (X * F, Y * F)-((X + 1) * F, (Y + 1) * F), RGB(255 - Field(X, Y).Steps / Max * 255, 255 - Field(X, Y).Steps / Max * 255, 255 - Field(X, Y).Steps / Max * 255), BF
      End If
      CurrentX = (X + 0.1) * F: CurrentY = Y * F
  '    Stop
      ForeColor = vbGreen
      FontSize = 11
      Print T
      
      If Field(X, Y).Block > 1 Or (X = XP And Y = YP) Then
        CurrentX = (X + 0.6) * F: CurrentY = Y * F + TextHeight("W") / 1.5
        ForeColor = vbBlue
        FontSize = 8
        Print FieldPoints(X, Y)
      End If
      
      Set TestCo = CreateCoordinates(X, Y)
      Ind = WayPointExists(TestCo)
      If Ind <> -1 Then
        If Ind = List2.ListIndex + 1 Then
          ForeColor = vbYellow
        Else
          ForeColor = vbGreen
        End If
        Line (X * F, Y * F)-((X + 1) * F, (Y + 1) * F), , BF
        If Check1.Value Then
          ForeColor = vbBlue
          CurrentX = (X + 0.1) * F: CurrentY = Y * F
          Print Ind
        End If
      End If
    End If
  Next
Next

ForeColor = vbBlack
For X = 0 To UBound(Field, 1) + 1
  Line (X * F, 0)-(X * F, (UBound(Field, 2) + 1) * F)
  CurrentY = UBound(Field, 2) * F + F
  CurrentX = X * F '* 0.9
  If X <= UBound(Field, 1) And IsMissing(X) Then Print X
Next
For X = 0 To UBound(Field, 2) + 1
  Line (0, X * F)-((UBound(Field, 1) + 1) * F, X * F)
  CurrentX = UBound(Field, 1) * F + F
  CurrentY = X * F '* 0.9
  If X <= UBound(Field, 2) And IsMissing(X) Then Print X
Next
ForeColor = RGB(200, 0, 0)
Circle (XP * F + F / 2, YP * F + F / 2), F / 3
ForeColor = RGB(150, 0, 0)
Circle (XP * F + F / 2, YP * F + F / 2), F / 3
End Sub
Private Function WayPointExists(Coordinate As Coordinate) As Integer
Dim Tel%
On Error GoTo Fout
WayPointExists = -1
MaakLijstWP
For Tel = 1 To WayPoints.Count
  If WayPoints(Tel).X = Coordinate.X And WayPoints(Tel).Y = Coordinate.Y Then
    WayPointExists = Tel
    Exit Function
  End If
Fout:
Next
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TestCo As Coordinate
X = Int(X / F)
Y = Int(Y / F)
If X > UBound(Field, 1) Or Y > UBound(Field, 2) Or X < 0 Or Y < 0 Then Exit Sub
If FindWay Or SetWayPoints Then GoTo Draw
If Button = 2 Then
  XP = X
  YP = Y
  Command1_Click
End If

If Button = 1 Then
  If List1.ListCount Then
    List1.Clear
    TekenPad
  End If
  If PrevX <> X Or PrevY <> Y Then
    If SetWayPoints = False Then
      Set TestCo = CreateCoordinates(CInt(X), CInt(Y))
      Field(X, Y).Block = IIf(Field(X, Y).Block = 1, 0, 1)
      If WayPointExists(TestCo) <> -1 Then
        WayPoints.Remove WayPointExists(TestCo)
        MaakLijstWP
      End If
    End If
  End If
  PrevX = X: PrevY = Y
End If

If Button <> 2 Then
Draw:
  
'  If Button <> 1 Or Command1.Enabled = False Then
    Teken PrevX, PrevY
'  End If
  If Field(X, Y).Block <> 1 Then
    Line (X * F, Y * F)-(X * F + F, Y * F + F), vbGreen, B
  Else
    Line (X * F, Y * F)-(X * F + F, Y * F + F), vbRed, B
  End If
End If
Label1 = "X=" + Str(X) + vbCrLf + "Y=" + Str(Y) + vbCrLf + "Steps=" + Str(Field(X, Y).Steps) + vbCrLf + "Total Points=" + Str(Field(X, Y).Points) ' + vbCrLf + "Block=" + Str(Field(X, Y).Block)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim P, T$, TestCo As New Coordinate
Dim StartCo As Coordinate, StopCo As Coordinate
Dim Tmr As Long

PrevX = -2
X = Int(X / F)
Y = Int(Y / F)
Command1_Click


Tmr = GetTickCount

Set StartCo = CreateCoordinates(XP, YP)
Set StopCo = CreateCoordinates(CInt(X), CInt(Y))
If FindWay And SetWayPoints = False Then
  If List2.ListCount Then
    If Check1.Value Then
      'in order
      Set P = FindPath(StartCo, StopCo, WayPoints, True)
    Else
      Set P = FindPath(StartCo, StopCo, WayPoints, False)
    End If
  Else
    'no waypoints set
    CreatePaths StartCo
    Set P = GetPath(StartCo, StopCo)
  End If
  List1.Clear
  If Not (P Is Nothing) Then
    For X = 1 To P.Count
      List1.AddItem Trim(Str(P(X).X)) + "," + Trim(Str(P(X).Y))
    Next
  End If
  Label2 = "Path Finding=" + Str(GetTickCount - Tmr) + vbCrLf
  TekenPad
  List1.ToolTipText = Str(List1.ListCount) + " steps"
End If

CreatePaths StartCo

If SetWayPoints = True Then
  Set TestCo = CreateCoordinates(CInt(X), CInt(Y))
  P = WayPointExists(TestCo)
  If P <> -1 Then
    WayPoints.Remove P
    Teken
  Else
    If Field(X, Y).Block <> e_Block Then
      WayPoints.Add TestCo
    End If
    Teken X, Y
  End If
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
scrX.Left = 0
scrX.Top = ScaleHeight - scrX.Height
scrX.Width = ScaleWidth - 250
scrY.Left = ScaleWidth - scrY.Width
scrY.Top = 0
scrY.Height = ScaleHeight - 250
''Command1.Left = Width - scrY.Width - Command1.Width - 250
''Command2.Left = Command1.Left
''List1.Left = Command1.Left
'Label1.Left = Width - scrY.Width - Label1.Width - 250
'Label2.Left = Width - scrY.Width - Label1.Width - 250
Frame1.Move Width - Frame1.Width - 600, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dim X%, Y%, D$
'D = Trim(Str(UBound(Field, 1))) + "," + Trim(Str(UBound(Field, 1))) + ";"
'D = D + Trim(Str(XP)) + "," + Trim(Str(YP)) + ";"
''D = D + Trim(Str(ToX)) + "," + Trim(Str(ToY)) + ";"
'For X = 0 To UBound(Field, 1)
'  For Y = 0 To UBound(Field, 2)
'    If Field(X, Y).Block = 1 Then
'      D = D + Trim(Str(Abs(Field(X, Y).Block)))
'    Else
'      D = D + "0"
'    End If
'  Next
'Next
'SaveSetting "BomberKid", "Pad Finder", "Field", D
End Sub

Private Sub List1_Click()
Dim T
Teken
TekenPad
T = Split(List1, ",")
Line (T(0) * F, T(1) * F)-((T(0) + 1) * F, (T(1) + 1) * F), vbBlue, BF
End Sub

Private Sub List2_Click()
Teken
End Sub

Private Sub scrX_Change()
Cls
ReDim Field(scrX, UBound(Field, 2))
CheckPos
SetPoints
Command1_Click
End Sub
Private Sub SetPoints()
Dim X%, Y%
ReDim FieldPoints(UBound(Field, 1), UBound(Field, 2))
For X = 0 To UBound(Field, 1)
  For Y = 0 To UBound(Field, 2)
    FieldPoints(X, Y) = Rnd * 2
  Next
Next
End Sub
Private Sub scrX_Scroll()
scrX_Change
End Sub

Private Sub scrY_Change()
Cls
ReDim Field(UBound(Field, 1), scrY)
CheckPos
SetPoints
Set WayPoints = New Collection
MaakLijstWP
Command1_Click
End Sub

Private Sub scrY_Scroll()
scrY_Change
End Sub
Private Sub CheckPos()
If XP > UBound(Field, 1) Then XP = UBound(Field, 1)
If YP > UBound(Field, 2) Then YP = UBound(Field, 2)
End Sub
