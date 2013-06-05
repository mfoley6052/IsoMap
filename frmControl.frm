VERSION 5.00
Begin VB.Form frmControl 
   Caption         =   "Form2"
   ClientHeight    =   2115
   ClientLeft      =   225
   ClientTop       =   1950
   ClientWidth     =   2445
   LinkTopic       =   "Form2"
   ScaleHeight     =   2115
   ScaleWidth      =   2445
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton Command3 
         Caption         =   "Clear Map"
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Load"
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load map"
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Dim counter As Integer
counter = 1
For X = 0 To 6
    For Y = 0 To 6
        counter = counter + 1
        Unload Form1.imgCube(counter)
        Unload Form1.shpTile(counter)
    Next Y
Next X
End Sub
Private Sub Command1_Click()
Call DrawIso("ISO")
End Sub

Private Sub Command2_Click()
Call DrawIso("Flat")
End Sub

