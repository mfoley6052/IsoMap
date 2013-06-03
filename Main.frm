VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   ScaleHeight     =   692
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   961
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   5160
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1920
      Top             =   3360
      Width           =   855
   End
   Begin VB.Shape shpTile 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   750
      Index           =   0
      Left            =   480
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Tile
    x As Integer
    y As Integer
    kind As Integer
    passable As Boolean
End Type
Dim FlatMap(6, 6) As Integer

Dim Map(6, 6) As Tile

Private Sub Command1_Click()
Call DrawIso
End Sub

Private Sub Form_Load()
For x = 0 To 6
    For y = 0 To 6
        If (x = 0 Or x = 6) And (y = 0 Or y = 6) Then
            FlatMap(x, y) = 1
        Else
            FlatMap(x, y) = 0
        End If
        Map(x, y).x = 50 * x
        Map(x, y).y = 50 * y
        Map(x, y).kind = FlatMap(x, y)
        Map(x, y).passable = True
    Next y
Next x

End Sub

Private Function Flat2Iso(xTile As Tile) As Tile
Dim cartx As Integer
Dim carty As Integer
cartx = xTile.x
carty = xTile.y
xTile.x = cartx - carty
xTile.y = (cartx + carty) / 2
Flat2Iso = xTile
End Function
Private Function Iso2Flat(xTile As Tile) As Tile
Dim cartx As Integer
Dim carty As Integer
cartx = xTile.x
carty = xTile.y
xTile.x = (2 * cartx + carty) / 2
xTile.y = (2 * cartx - carty) / 2
Iso2Flat = xTile
End Function

Private Sub DrawIso()
Dim counter As Integer
counter = 1
For x = 0 To 6
    For y = 0 To 6
        counter = counter + 1
        Load shpTile(counter)
        If Map(x, y).kind = 0 Then
            shpTile(counter).FillColor = vbGreen
        ElseIf Map(x, y).kind = 1 Then
            shpTile(counter).FillColor = vbRed
        End If
        shpTile(counter).Left = Map(x, y).x
        shpTile(counter).Top = Map(x, y).y
        shpTile(counter).Visible = True
    Next y
Next x
End Sub