Attribute VB_Name = "Module1"
Option Explicit
Private Visited() As Boolean
Public Field() As t_Vld
Public FieldPoints() As Integer
Public Type t_Vld
  Block As enum_Block
  Steps As Integer
  Points As Integer
End Type
Public Enum enum_Block
  e_Empty = 0
  e_Block = 1
  e_Up = 2
  e_Right = 3
  e_Down = 4
  e_Left = 5
End Enum
Private SomeThingChanged As Boolean

Public Function CreateCoordinates(XPos As Integer, YPos As Integer) As Coordinate
Set CreateCoordinates = New Coordinate
CreateCoordinates.X = XPos
CreateCoordinates.Y = YPos
End Function
Private Function IsInRange(Co As Coordinate) As Boolean
If Co.X < 0 Or Co.X > UBound(Field, 1) Then
  Exit Function
End If
If Co.Y < 0 Or Co.Y > UBound(Field, 2) Then
  Exit Function
End If
IsInRange = True
End Function
Function FindPath(ByVal StartCo As Coordinate, ByVal StopCo As Coordinate, ByVal WayPoints As Collection, ByVal InOrder As Boolean) As Collection
Dim X%, Y%
Dim CurCo As Coordinate
Dim TempPath As Collection
Dim Shortest As Integer, ShortestCO As Coordinate, ShortestIndex%
Dim WP As New Collection
'find a path with waypoints:
' in an order
' no order-->search for the nearest waypoint
If WayPoints Is Nothing Then Exit Function
If IsInRange(StartCo) = False Or IsInRange(StopCo) = False Then Exit Function
For X = 1 To WayPoints.Count
  WP.Add WayPoints(X)
Next
Set FindPath = New Collection
Set CurCo = StartCo
If InOrder Then
  For X = 1 To WP.Count
    CreatePaths CurCo
    Set TempPath = GetPath(CurCo, WP(X))
    If Not (TempPath Is Nothing) Then
      For Y = 1 To TempPath.Count
        FindPath.Add TempPath(Y)
      Next
      Set CurCo = WP(X)
    End If
  Next
  CreatePaths CurCo
  Set TempPath = GetPath(CurCo, StopCo)
  If Not (TempPath Is Nothing) Then
    For Y = 1 To TempPath.Count
      FindPath.Add TempPath(Y)
    Next
  End If
Else
  Do
    CreatePaths CurCo
    Shortest = 600
    For X = 1 To WP.Count
      If Field(WP(X).X, WP(X).Y).Steps < Shortest Then
        Shortest = Field(WP(X).X, WP(X).Y).Steps
        Set ShortestCO = WP(X)
        ShortestIndex = X
      End If
    Next
    Set TempPath = GetPath(CurCo, ShortestCO)
    If Not (TempPath Is Nothing) Then
      For Y = 1 To TempPath.Count
        FindPath.Add TempPath(Y)
      Next
      Set CurCo = ShortestCO
    End If
    WP.Remove ShortestIndex
  Loop Until WP.Count = 0
  Set TempPath = GetPath(CurCo, StopCo)
  For Y = 1 To TempPath.Count
    FindPath.Add TempPath(Y)
  Next
End If
For X = FindPath.Count To 2 Step -1
  If FindPath.Item(X - 1).X = FindPath.Item(X).X And FindPath.Item(X - 1).Y = FindPath.Item(X).Y Then
    FindPath.Remove X
    X = X - 1
  End If
Next
End Function
Sub CreatePaths(ByVal StartCo As Coordinate)
Dim X%, Y%
'initiate the field
ReDim Visited(UBound(Field, 1), UBound(Field, 2))
For X = 0 To UBound(Field, 1)
  For Y = 0 To UBound(Field, 2)
    Field(X, Y).Steps = 0
    Field(X, Y).Points = 0
  Next
Next

'start looking up
Up StartCo.X, StartCo.Y, 0
Down StartCo.X, StartCo.Y + 1, 0

'correct the paths, so it'll take the shortest path
If Field(StartCo.X, StartCo.Y).Block <> e_Block Then Field(StartCo.X, StartCo.Y).Block = e_Empty
Field(StartCo.X, StartCo.Y).Points = FieldPoints(StartCo.X, StartCo.Y)
CorrectPaths ByVal StartCo.X, ByVal StartCo.Y
End Sub
Private Sub CorrectPaths(ByVal StartX%, ByVal StartY%)
'correct the paths, so the shortest path is taken
'do it until there is nothing to correct
Do
  SomeThingChanged = False
  ReDim Visited(UBound(Field, 1), UBound(Field, 2))
  Up StartX, StartY, 0, True
  ReDim Visited(UBound(Field, 1), UBound(Field, 2))
  Down StartX, StartY + 1, 0, True
Loop Until SomeThingChanged = False
End Sub
Public Function GetPath(ByVal StartCo As Coordinate, ByVal StopCo As Coordinate) As Collection
'get the path from a point to another
Set GetPath = New Collection
Dim TempPath As New Collection
Dim X%, Skip As Boolean, Co As Coordinate
Dim ToX%, ToY%
ToX = StopCo.X: ToY = StopCo.Y
CreatePaths StartCo
If Field(ToX, ToY).Block = e_Block Then Exit Function
Do
  Set Co = CreateCoordinates(ToX, ToY)
  TempPath.Add Co
  Skip = False
  If Field(ToX, ToY).Block = e_Down Then
    ToY = ToY + 1
    Skip = True
  End If
  If Field(ToX, ToY).Block = e_Up And Skip = False Then
    ToY = ToY - 1
    Skip = True
  End If
  If Field(ToX, ToY).Block = e_Right And Skip = False Then
    ToX = ToX + 1
    Skip = True
  End If
  If Field(ToX, ToY).Block = e_Left And Skip = False Then
    ToX = ToX - 1
    Skip = True
  End If
  If Skip = False Then
    Set GetPath = Nothing
    Exit Function
  End If
Loop Until ToX = StartCo.X And ToY = StartCo.Y
Set Co = CreateCoordinates(ToX, ToY)
TempPath.Add Co
For X = TempPath.Count To 1 Step -1
  GetPath.Add TempPath(X)
Next
End Function
Private Sub Up(ByVal X%, ByVal Y%, ByVal Steps%, Optional CheckEnvironment As Boolean)
Dim N%
If X < 0 Then Exit Sub
For N = Y To 0 Step -1
  'filled block
  If Field(X, N).Block = e_Block Then Exit For
  'already been
  If Visited(X, N) Then Exit For
  'we have visited this block, so don't do that enymore
  Visited(X, N) = True
  'do we have to create a path, or correct it?
  If CheckEnvironment = False Then
    Field(X, N).Block = e_Down
    Field(X, N).Steps = Steps
  Else
    Environment X, N
  End If
  'go left, and right
  Left X - 1, N, Steps + 1, CheckEnvironment
  Right X + 1, N, Steps + 1, CheckEnvironment
  'add a step to the total
  Steps = Steps + 1
Next
End Sub
Private Sub Down(ByVal X%, ByVal Y%, ByVal Steps%, Optional CheckEnvironment As Boolean)
Dim N%
If X < 0 Then Exit Sub
If Field(X, Y - 1).Block = 1 Then Exit Sub
For N = Y To UBound(Field, 2)
  If Field(X, N).Block = 1 Then Exit For
  If Visited(X, N) Then Exit For
  Visited(X, N) = True
  If CheckEnvironment = False Then
    Field(X, N).Block = e_Up
    Field(X, N).Steps = Steps
  Else
    Environment X, N
  End If
  Left X - 1, N, Steps + 1, CheckEnvironment
  Right X + 1, N, Steps + 1, CheckEnvironment
  Steps = Steps + 1
Next
End Sub
Private Sub Left(ByVal X%, ByVal Y%, ByVal Steps%, Optional CheckEnvironment As Boolean)
Dim N%
For N = X To 0 Step -1
  If Field(N, Y).Block = 1 Then Exit For
  If Visited(N, Y) Then Exit For
  Visited(N, Y) = True
  If CheckEnvironment = False Then
    Field(N, Y).Block = e_Right
    Field(N, Y).Steps = Steps
  Else
    Environment N, Y
  End If
  Up N, Y - 1, Steps + 1, CheckEnvironment
  Down N, Y + 1, Steps + 1, CheckEnvironment
  Steps = Steps + 1
Next
End Sub
Private Sub Right(ByVal X%, ByVal Y%, ByVal Steps%, Optional CheckEnvironment As Boolean)
Dim N%
For N = X To UBound(Field, 1)
  If Field(N, Y).Block = 1 Then Exit For
  If Visited(N, Y) Then Exit For
  Visited(N, Y) = True
  If CheckEnvironment = False Then
    Field(N, Y).Block = e_Left
    Field(N, Y).Steps = Steps
  Else
    Environment N, Y
  End If
  Up N, Y - 1, Steps + 1, CheckEnvironment
  Down N, Y + 1, Steps + 1, CheckEnvironment
  Steps = Steps + 1
Next
End Sub
Private Sub Environment(X%, Y%)
Dim S%
'check the surroundings, and if there is a block with a lower
'number of steps, go to that direction

S = Field(X, Y).Steps
If X > 0 Then
  If Field(X - 1, Y).Steps > S + 1 Then
    Field(X - 1, Y).Block = e_Right
    Field(X - 1, Y).Steps = Field(X, Y).Steps + 1
    SomeThingChanged = True
  End If
  If Field(X - 1, Y).Block = e_Right Then
    Field(X - 1, Y).Points = Field(X, Y).Points + FieldPoints(X - 1, Y)
  End If
End If
If Y > 0 Then
  If Field(X, Y - 1).Steps > S + 1 Then
    Field(X, Y - 1).Block = e_Down
    Field(X, Y - 1).Steps = Field(X, Y).Steps + 1
    SomeThingChanged = True
  End If
  If Field(X, Y - 1).Block = e_Down Then
    Field(X, Y - 1).Points = Field(X, Y).Points + FieldPoints(X, Y - 1)
  End If
End If
If X < UBound(Field, 1) Then
  If Field(X + 1, Y).Steps > S + 1 Then
    Field(X + 1, Y).Block = e_Left
    Field(X + 1, Y).Steps = Field(X, Y).Steps + 1
    SomeThingChanged = True
  End If
  If Field(X + 1, Y).Block = e_Left Then
    Field(X + 1, Y).Points = Field(X, Y).Points + FieldPoints(X + 1, Y)
  End If
End If
If Y < UBound(Field, 2) Then
  If Field(X, Y + 1).Steps > S + 1 Then
    Field(X, Y + 1).Block = e_Up
    Field(X, Y + 1).Steps = Field(X, Y).Steps + 1
    SomeThingChanged = True
  End If
  If Field(X, Y + 1).Block = e_Up Then
    Field(X, Y + 1).Points = Field(X, Y).Points + FieldPoints(X, Y + 1)
  End If
End If
End Sub
