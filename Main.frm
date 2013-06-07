VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Form1"
   ClientHeight    =   10380
   ClientLeft      =   4365
   ClientTop       =   2280
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   ScaleHeight     =   692
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   961
   Begin VB.Timer tmrAi 
      Interval        =   750
      Left            =   480
      Top             =   4560
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      Caption         =   "00000000"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12960
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Shape shpPlayer 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   1
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape shpTile 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   0
      Left            =   480
      Shape           =   1  'Square
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Shape shpPlayer 
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   0
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgCube 
      Height          =   750
      Index           =   0
      Left            =   11760
      Picture         =   "Main.frx":0000
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NodeCount As Integer
Dim NodeMap(6, 6) As Tile
Dim numSteps As Integer
Dim startTile As Tile
Dim targetTile As Tile
Dim curTile As Tile
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
    If curx(0) > 0 Then
        If Map(curx(0) - 1, cury(0)).passable Then
            lblScore.Caption = Format(Val(lblScore.Caption) + 10, "00000000")
            Map(curx(0), cury(0)).used = False
            curx(0) = curx(0) - 1
            Map(curx(0), cury(0)).used = True
        End If
    End If
ElseIf KeyCode = vbKeyRight Then
    If curx(0) < 6 Then
        If Map(curx(0) + 1, cury(0)).passable Then
            lblScore.Caption = Format(Val(lblScore.Caption) + 10, "00000000")
            Map(curx(0), cury(0)).used = False
            curx(0) = curx(0) + 1
            Map(curx(0), cury(0)).used = True
        End If
    End If
ElseIf KeyCode = vbKeyDown Then
    If cury(0) < 6 Then
        If Map(curx(0), cury(0) + 1).passable Then
            lblScore.Caption = Format(Val(lblScore.Caption) + 10, "00000000")
            Map(curx(0), cury(0)).used = False
            cury(0) = cury(0) + 1
            Map(curx(0), cury(0)).used = True
        End If
    End If
ElseIf KeyCode = vbKeyUp Then
    If cury(0) > 0 Then
        If Map(curx(0), cury(0) - 1).passable Then
            lblScore.Caption = Format(Val(lblScore.Caption) + 10, "00000000")
            Map(curx(0), cury(0)).used = False
            cury(0) = cury(0) - 1
            Map(curx(0), cury(0)).used = True
        End If
    End If
End If
shpPlayer(0).Left = Map(curx(0), cury(0)).x + 15
shpPlayer(0).Top = Map(curx(0), cury(0)).y
shpPlayer(0).Visible = True
If curx(0) = curx(1) And cury(0) = cury(1) Then
    If Val(lblScore.Caption) >= 150 Then
        lblScore.Caption = Format(Val(lblScore.Caption) - 150, "00000000")
    Else
        lblScore.Caption = Format(0, "00000000")
    End If
End If
End Sub

Private Sub Form_Load()
For x = 0 To 6
    For y = 0 To 6
        If x = 0 Or x = 6 Or y = 0 Or y = 6 Then
            FlatMap(x, y) = 1
        Else
            FlatMap(x, y) = 0
        End If
        Map(x, y).x = 50 * x + 350
        Map(x, y).y = 50 * y
        Map(x, y).kind = FlatMap(x, y)
        Map(x, y).passable = True
    Next y
Next x
curx(0) = 2
cury(0) = 2
curx(1) = 5
cury(1) = 3
Map(2, 2).used = True
Map(5, 3).used = True
Call DrawIso("ISO")
End Sub
Private Sub GetStep(ByVal stepNum As Integer)
Do Until curx(1) = path(stepNum).x And cury(1) = path(stepNum).y
    Map(curx(1), cury(1)).used = False
    If stepNum > LBound(path) Then
        If path(stepNum).x > curx(1) Then
            curx(1) = curx(1) + 1
        ElseIf path(stepNum).x < curx(1) Then
            curx(1) = curx(1) - 1
        End If
        If path(stepNum).y > cury(1) Then
            cury(1) = cury(1) + 1
        ElseIf path(stepNum).y < cury(1) Then
            cury(1) = cury(1) - 1
        End If
    End If
    Map(curx(1), cury(1)).used = True
    shpPlayer(1).Left = Map(curx(1), cury(1)).x + 15
    shpPlayer(1).Top = Map(curx(1), cury(1)).y
    shpPlayer(1).Visible = True
    If curx(0) = curx(1) And cury(0) = cury(1) Then
        If Val(lblScore.Caption) >= 150 Then
            lblScore.Caption = Format(Val(lblScore.Caption) - 150, "00000000")
        Else
            lblScore.Caption = Format(0, "00000000")
        End If
    End If
Loop
End Sub
Private Sub imgCube_Click(index As Integer)
Unload imgCube(index)
Map(info(index).x, info(index).y).passable = False
info(index).passable = False
End Sub
Private Sub GetTile()
Dim smallestF As Tile
Dim onList As Boolean
smallestF = Map(curx(1), cury(1))
Do Until onList
    ReDim Preserve closedList(numSteps) As Tile
    ReDim openList(numSteps) As Tile
    For x = LBound(closedList) To UBound(closedList)
        If closedList(x).x = targetTile.x And closedList(x).y = targetTile.y Then
            onList = True
        End If
    Next x
    Call GetAdj(smallestF)
    Dim tempCount As Integer
    For x = LBound(openList) To UBound(openList) - 1
        If openList(x).F < openList(x + 1).F Then
            smallestF = openList(x)
        End If
    Next x
    closedList(numSteps) = smallestF
    For x = LBound(openList) To UBound(openList)
        If openList(x).x = smallestF.x And openList(x).y = smallestF.y Then
            ReDim tempList(UBound(openList) - 1) As Tile
            For y = LBound(openList) To UBound(openList)
                If openList(y).x <> smallestF.x And openList(y).y <> smallestF.y Then
                    tempList(tempCount) = openList(y)
                    tempCount = tempCount + 1
                End If
            Next y
            Erase openList
            ReDim openList(UBound(tempList)) As Tile
            openList = tempList
            Erase tempList
        End If
    Next x
    numSteps = numSteps + 1
Loop
ReDim Preserve path(UBound(closedList)) As Tile
path = closedList
End Sub
Private Sub GetAdj(aTile As Tile)
Dim onListC As Boolean
Dim onListO As Boolean
curTile = Map(aTile.tilex, aTile.tiley)
targetTile = Map(curx(0), cury(0))
If NodeCount = 0 Then
    startTile = Map(curx(1), cury(1))
End If
For x = LBound(closedList) To UBound(closedList)
    If closedList(x).x = Map(curTile.tilex + 1, curTile.tiley).x And closedList(x).y = Map(curTile.tilex + 1, curTile.tiley).y Then
        onListC = True
    End If
Next x
For x = LBound(openList) To UBound(openList)
    If openList(x).x = Map(curTile.tilex + 1, curTile.tiley).x And openList(x).y = Map(curTile.tilex + 1, curTile.tiley).y Then
        onListO = True
    End If
Next x
If Map(curTile.tilex + 1, curTile.tiley).passable And onListC = False And onListO = False Then
    ReDim Preserve openList(NodeCount) As Tile
    NodeMap(curTile.tilex + 1, curTile.tiley).G = Abs(curTile.tilex + 1 - startTile.tilex) + Abs(curTile.tiley - startTile.tiley)
    NodeMap(curTile.tilex + 1, curTile.tiley).H = Abs(curTile.tilex + 1 - targetTile.tilex) + Abs(curTile.tiley - targetTile.tiley)
    NodeMap(curTile.tilex + 1, curTile.tiley).F = NodeMap(curTile.tilex + 1, curTile.tiley).G + NodeMap(curTile.tilex + 1, curTile.tiley).H
    openList(NodeCount) = NodeMap(curTile.tilex + 1, curTile.tiley)
    NodeCount = NodeCount + 1
End If
onListC = False
onListO = False
'Continue

For x = LBound(closedList) To UBound(closedList)
    If closedList(x).x = Map(curTile.tilex - 1, curTile.tiley).x And closedList(x).y = Map(curTile.tilex - 1, curTile.tiley).y Then
        onListC = True
    End If
Next x
If Map(curTile.tilex - 1, curTile.tiley).passable And onList = False Then
    ReDim Preserve openList(NodeCount) As Tile
    NodeMap(curTile.tilex - 1, curTile.tiley).G = Abs(curTile.tilex - startTile.tilex) + Abs(curTile.tiley - startTile.tiley)
    NodeMap(curTile.tilex - 1, curTile.tiley).H = Abs(curTile.tilex - targetTile.tilex) + Abs(curTile.tiley - targetTile.tiley)
    NodeMap(curTile.tilex - 1, curTile.tiley).F = NodeMap(curTile.tilex - 1, curTile.tiley).G + NodeMap(curTile.tilex - 1, curTile.tiley).H
    openList(NodeCount) = NodeMap(curTile.tilex - 1, curTile.tiley)
    NodeCount = NodeCount + 1
End If
onList = False
For x = LBound(closedList) To UBound(closedList)
    If closedList(x).x = Map(curTile.tilex, curTile.tiley + 1).x And closedList(x).y = Map(curTile.tilex, curTile.tiley + 1).y Then
        onList = True
    End If
Next x
If Map(curTile.tilex, curTile.tiley + 1).passable And onList = False Then
    ReDim Preserve openList(NodeCount) As Tile
    NodeMap(curTile.tilex, curTile.tiley + 1).G = Abs(curTile.tilex - startTile.tilex) + Abs(curTile.tiley - startTile.tiley)
    NodeMap(curTile.tilex, curTile.tiley + 1).H = Abs(curTile.tilex - targetTile.tilex) + Abs(curTile.tiley - targetTile.tiley)
    NodeMap(curTile.tilex, curTile.tiley + 1).F = NodeMap(curTile.tilex, curTile.tiley + 1).G + NodeMap(curTile.tilex, curTile.tiley + 1).H
    openList(NodeCount) = NodeMap(curTile.tilex, curTile.tiley + 1)
    NodeCount = NodeCount + 1
End If
onList = False
For x = LBound(closedList) To UBound(closedList)
    If closedList(x).x = Map(curTile.tilex, curTile.tiley - 1).x And closedList(x).y = Map(curTile.tilex, curTile.tiley - 1).y Then
        onList = True
    End If
Next x
If Map(curTile.tilex, curTile.tiley - 1).passable And onList = False Then
    ReDim Preserve openList(NodeCount) As Tile
    NodeMap(curTile.tilex, curTile.tiley - 1).G = Abs(curTile.tilex - startTile.tilex) + Abs(curTile.tiley - startTile.tiley)
    NodeMap(curTile.tilex, curTile.tiley - 1).H = Abs(curTile.tilex - targetTile.tilex) + Abs(curTile.tiley - targetTile.tiley)
    NodeMap(curTile.tilex, curTile.tiley - 1).F = NodeMap(curTile.tilex, curTile.tiley - 1).G + NodeMap(curTile.tilex, curTile.tiley - 1).H
    openList(NodeCount) = NodeMap(curTile.tilex, curTile.tiley - 1)
    NodeCount = NodeCount + 1
End If
End Sub

Private Sub tmrAi_Timer()
'Dim temp As Integer
'temp = Int(Rnd() * 2)
'Map(curx(1), cury(1)).used = False
'
'If temp = 0 Then
'    If curx(1) >= curx(0) Then
'        If Map(curx(1) - 1, cury(1)).passable Then
'            curx(1) = curx(1) - 1
'        End If
'    Else
'        If Map(curx(1) + 1, cury(1)).passable Then
'            curx(1) = curx(1) + 1
'        End If
'    End If
'Else
'    If cury(1) >= cury(0) Then
'        If Map(curx(1), cury(1) - 1).passable Then
'           cury(1) = cury(1) - 1
'        End If
'    Else
'        If Map(curx(1), cury(1) + 1).passable Then
'            cury(1) = cury(1) + 1
'        End If
'    End If
'End If
'Map(curx(1), cury(1)).used = True
'shpPlayer(1).Left = Map(curx(1), cury(1)).x + 15
'shpPlayer(1).Top = Map(curx(1), cury(1)).y
'shpPlayer(1).Visible = True
'If curx(0) = curx(1) And cury(0) = cury(1) Then
'    If Val(lblScore.Caption) >= 150 Then
'        lblScore.Caption = Format(Val(lblScore.Caption) - 150, "00000000")
'    Else
'        lblScore.Caption = Format(0, "00000000")
'    End If
'End If
Call GetTile
For x = UBound(path) To LBound(path) Step -1
    Call GetStep(x)
Next x
End Sub
