Attribute VB_Name = "modFunct"
Public FlatMap(6, 6) As Integer
Public Map(6, 6) As Tile
Public curx(1) As Integer
Public cury(1) As Integer
Public Type Tile
    X As Integer
    Y As Integer
    kind As Integer
    used As Boolean
    passable As Boolean
End Type
Public Sub DrawIso(ByVal whichType As String)
Dim counter As Integer
With Main
counter = 1
For X = 0 To 6
    For Y = 0 To 6
        counter = counter + 1
        Load .shpTile(counter)
        Load .imgCube(counter)
        .imgCube(counter).Visible = True
        If Map(X, Y).kind = 0 Then
            .imgCube(counter).Picture = LoadPicture(App.Path & "/gblock.gif")
            .shpTile(counter).FillColor = vbGreen
        ElseIf Map(X, Y).kind = 1 Then
            '.imgCube(counter).Picture = LoadPicture(App.Path & "/block.gif")
            .shpTile(counter).FillColor = vbRed
        End If
        If whichType = "ISO" Then
            .imgCube(counter).Left = Flat2Iso(Map(X, Y)).X
            .imgCube(counter).Top = Map(X, Y).Y
        Else
            .imgCube(counter).Left = Iso2Flat(Map(X, Y)).X
            .imgCube(counter).Top = Map(X, Y).Y
        End If
        
'        .shpTile(counter).Left = Map(X, Y).X
'        .shpTile(counter).Top = Map(X, Y).Y
'        .shpTile(counter).Visible = True
    Next Y
Next X
For X = .shpPlayer.Lbound To .shpPlayer.Ubound
    .shpPlayer(X).Left = Map(curx(X), cury(X)).X + 15
    .shpPlayer(X).Top = Map(curx(X), cury(X)).Y
    .shpPlayer(X).Visible = True
Next X
End With
End Sub
Public Function Flat2Iso(xTile As Tile) As Tile
Dim cartx As Integer
Dim carty As Integer
cartx = xTile.X
carty = xTile.Y
xTile.X = cartx - carty
xTile.Y = (cartx + carty) / 2
Flat2Iso = xTile
End Function
Public Function Iso2Flat(xTile As Tile) As Tile
Dim cartx As Integer
Dim carty As Integer
cartx = xTile.X
carty = xTile.Y
xTile.X = (2 * cartx + carty) / 2
xTile.Y = (2 * cartx - carty) / 2
Iso2Flat = xTile
End Function
