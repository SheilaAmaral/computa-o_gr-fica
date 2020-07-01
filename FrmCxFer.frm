VERSION 5.00
Begin VB.Form frmCxFerr 
   BackColor       =   &H80000004&
   Caption         =   "PhotoX"
   ClientHeight    =   4710
   ClientLeft      =   1050
   ClientTop       =   1620
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleMode       =   0  'User
   ScaleWidth      =   1725
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   360
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   13
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   12
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdCxF 
      Height          =   615
      Index           =   11
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2985
      Width           =   615
   End
   Begin VB.CommandButton cmdCxF 
      Height          =   615
      Index           =   10
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2985
      Width           =   615
   End
   Begin VB.CommandButton cmdCxF 
      Height          =   615
      Index           =   9
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdCxF 
      Height          =   615
      Index           =   8
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdCxF 
      Height          =   615
      Index           =   7
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdCxF 
      Height          =   615
      Index           =   6
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdCxF 
      Height          =   615
      Index           =   5
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdCxF 
      Height          =   615
      Index           =   4
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdCxF 
      Height          =   615
      Index           =   3
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdCxF 
      Height          =   615
      Index           =   2
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdCxF 
      Height          =   615
      Index           =   1
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdCxF 
      Height          =   615
      Index           =   0
      Left            =   0
      Picture         =   "FrmCxFer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmCxFerr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCxF_Click(Index As Integer)
glbFerrSelec = Index
Select Case Index
Case 0
   ' MDIForm1.StatusBar1.SimpleText = "Linha Reta"
Case 1
   '  MDIForm1.StatusBar1.SimpleText = "Polígonos"
Case 2
    'MDIForm1.StatusBar1.SimpleText = "Formas"
Case 3
    'MDIForm1.StatusBar1.SimpleText = "Desenho Livre"
Case 4
    'MDIForm1.StatusBar1.SimpleText = "Círculo"
Case 5
    'MDIForm1.StatusBar1.SimpleText = "Elipse"
Case 6
    'MDIForm1.StatusBar1.SimpleText = "Borracha"
Case 7
    'MDIForm1.StatusBar1.SimpleText = "Curvas"
Case 8
    'MDIForm1.StatusBar1.SimpleText = "Fractais"
Case 9
    'MDIForm1.StatusBar1.SimpleText = "Texto"
    'vIf glbFerrSelec = 9 Then
     '   frmTxt.Show
    'End If
Case 10
    'MDIForm1.StatusBar1.SimpleText = "Hexágono"
Case Else
    'MDIForm1.StatusBar1.SimpleText = "Recortar"
End Select

End Sub

Private Sub Form_Load()

Me.Caption = "PhotoX"
Me.Height = 5220
Me.Width = 1350
Me.Top = 0
Me.Left = 0

End Sub

Private Sub picColor_Click(Index As Integer)
    Select Case Index
    Case 0
        glbFerrSelec = 12
    Case 1
        glbFerrSelec = 13
    End Select
End Sub


Private Sub picColor_DblClick(Index As Integer)
    frmPaleta.Show
End Sub
