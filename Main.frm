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

Private Sub imgCube_Click(index As Integer)
Unload imgCube(index)
Map(info(index).x, info(index).y).passable = False
info(index).passable = False
End Sub

Private Sub tmrAi_Timer()
Dim temp As Integer
temp = Int(Rnd() * 2)
Map(curx(1), cury(1)).used = False

If temp = 0 Then
    If curx(1) >= curx(0) Then
        If Map(curx(1) - 1, cury(1)).passable Then
            curx(1) = curx(1) - 1
        End If
    Else
        If Map(curx(1) + 1, cury(1)).passable Then
            curx(1) = curx(1) + 1
        End If
    End If
Else
    If cury(1) >= cury(0) Then
        If Map(curx(1), cury(1) - 1).passable Then
           cury(1) = cury(1) - 1
        End If
    Else
        If Map(curx(1), cury(1) + 1).passable Then
            cury(1) = cury(1) + 1
        End If
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
End Sub
