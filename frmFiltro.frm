VERSION 5.00
Begin VB.Form frmFiltro 
   Caption         =   "Filtro"
   ClientHeight    =   1140
   ClientLeft      =   1755
   ClientTop       =   1545
   ClientWidth     =   2355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   2355
   Begin VB.CommandButton cmdFiltroOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtFiltro 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nº DE LADOS"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFiltroOK_Click()
glbFiltro = Val(txtFiltro.Text)
Unload Me
End Sub

Private Sub Form_Load()
Me.Left = 2115
Me.Top = 2145
Me.Width = 2475
Me.Height = 1545
End Sub
