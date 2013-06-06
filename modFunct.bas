Attribute VB_Name = "modFunct"
Public FlatMap(6, 6) As Integer
Public Map(6, 6) As Tile
Public curx(1) As Integer
Public cury(1) As Integer
Public info() As Tile
Public Type Tile
    x As Integer
    y As Integer
    kind As Integer
    index As Integer
    used As Boolean
    passable As Boolean
End Type
Public Sub DrawIso(ByVal whichType As String)
Dim counter As Integer
With Main
counter = 1
For x = 0 To 6
    For y = 0 To 6
        counter = counter + 1
        ReDim Preserve info(counter) As Tile
        info(counter).index = counter
        info(counter).x = x
        info(counter).y = y
        info(counter).passable = True
        Load .shpTile(counter)
        Load .imgCube(counter)
        .imgCube(counter).Visible = True
        If Map(x, y).kind = 0 Then
            .imgCube(counter).Picture = LoadPicture(App.Path & "/gblock.gif")
            .shpTile(counter).FillColor = vbGreen
            Map(x, y).passable = True
        ElseIf Map(x, y).kind = 1 Then
            '.imgCube(counter).Picture = LoadPicture(App.Path & "/block.gif")
            .shpTile(counter).FillColor = vbRed
            Map(x, y).passable = False
        End If
        If whichType = "ISO" Then
            .imgCube(counter).Left = Flat2Iso(Map(x, y)).x
            .imgCube(counter).Top = Map(x, y).y
        Else
            .imgCube(counter).Left = Iso2Flat(Map(x, y)).x
            .imgCube(counter).Top = Map(x, y).y
        End If
        
        '.shpTile(counter).Left = Map(X, Y).X
        '.shpTile(counter).Top = Map(X, Y).Y
        '.shpTile(counter).Visible = True
    Next y
Next x
For x = .shpPlayer.Lbound To .shpPlayer.Ubound
    .shpPlayer(x).Left = Map(curx(x), cury(x)).x + 15
    .shpPlayer(x).Top = Map(curx(x), cury(x)).y
    .shpPlayer(x).Visible = True
Next x
End With
End Sub
Public Function Flat2Iso(xTile As Tile) As Tile
Dim cartx As Integer
Dim carty As Integer
cartx = xTile.x
carty = xTile.y
xTile.x = cartx - carty
xTile.y = (cartx + carty) / 2
Flat2Iso = xTile
End Function
Public Function Iso2Flat(xTile As Tile) As Tile
Dim cartx As Integer
Dim carty As Integer
cartx = xTile.x
carty = xTile.y
xTile.x = (2 * cartx + carty) / 2
xTile.y = (2 * cartx - carty) / 2
Iso2Flat = xTile
End Function
