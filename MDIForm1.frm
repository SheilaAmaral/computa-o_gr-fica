VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000B&
   Caption         =   "PhotoX"
   ClientHeight    =   5610
   ClientLeft      =   3165
   ClientTop       =   2715
   ClientWidth     =   6585
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuImag 
      Caption         =   "&Image"
   End
   Begin VB.Menu mnuFiltro 
      Caption         =   "Fi&lter"
      Begin VB.Menu mnuFlt 
         Caption         =   "Linha Reta"
         Index           =   0
         Begin VB.Menu mnuReta 
            Caption         =   "Usar DDA"
            Index           =   0
         End
         Begin VB.Menu mnuReta 
            Caption         =   "Usar Bresenham"
            Index           =   1
         End
         Begin VB.Menu mnuReta 
            Caption         =   "Usar Método do VB"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFlt 
         Caption         =   "Poligonos"
         Index           =   1
         Begin VB.Menu mnuPolig 
            Caption         =   "Triângulo"
            Index           =   0
         End
         Begin VB.Menu mnuPolig 
            Caption         =   "Retângulo"
            Index           =   1
         End
         Begin VB.Menu mnuPolig 
            Caption         =   "Pentagono"
            Index           =   2
         End
         Begin VB.Menu mnuPolig 
            Caption         =   "Estrela"
            Index           =   3
         End
         Begin VB.Menu mnuPolig 
            Caption         =   "Regular"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFlt 
         Caption         =   "Círculo"
         Index           =   3
         Begin VB.Menu mnuCurv 
            Caption         =   "Usar Trigonometria"
            Index           =   0
         End
         Begin VB.Menu mnuCurv 
            Caption         =   "Usar Incremental"
            Index           =   1
         End
         Begin VB.Menu mnuCurv 
            Caption         =   "Usar Método da Linguagem"
            Index           =   2
         End
         Begin VB.Menu mnuCurv 
            Caption         =   "Arcos"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuExib 
      Caption         =   "&View"
   End
   Begin VB.Menu mnuWin 
      Caption         =   "&Window"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Me.Caption = "PhotoEu"
Me.Left = 1425
Me.Top = 1110
Me.Width = 9750
Me.Height = 7590
glbProxPonto = False
frmDesenho.Show
frmCxFerr.Show
End Sub

Private Sub mnuCurv_Click(Index As Integer)
    glbMetodo = Index + 8
End Sub

Private Sub mnuflt_Click(Index As Integer)
glbMetodo = Index
End Sub

Private Sub mnuPolig_Click(Index As Integer)
glbMetodo = Index + 3
If glbMetodo = 7 Then
    frmFiltro.Show
End If
End Sub
