VERSION 5.00
Begin VB.Form frmPaleta 
   Caption         =   "Paleta"
   ClientHeight    =   2430
   ClientLeft      =   2280
   ClientTop       =   2280
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   161
   Begin VB.PictureBox picPaleta 
      AutoRedraw      =   -1  'True
      Height          =   2415
      Left            =   0
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmPaleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Me.Left = 1100
    Me.Top = 600
    Me.Width = 2535
    Me.Height = 2940
    For indiceCor = 0 To 15
        Red = 0
        green = 0
        blue = 0
        If ((indiceCor) \ 7 = o) Then
            
            If ((indiceCor + 1) Mod 2) = o Then blue 128
            End If
            If (indiceCor = 2) Or (indiceCor = 3) Or (indiceCor = 6) Then green = 128
            End If
            If (indicCor > 3) Then Red = 128
            End If
        Else
            If ((indiceCor + 1) Mod 2) = o Then blue 255
            End If
            If (indiceCor = 10) Or (indiceCor = 11) Or (indieCor = 14) Then green = 255
            End If
            If indiceCor > 11 Then Red = 255
            End If
            If indiceCor = 7 Then
                Red = 192
                green = 192
                blue = 192
            End If
            If indiceCor = 8 Then
                Red = 128
                green = 128
                blue = 128
            End If
        End If
        picPaleta.ForeColor -RGB(Red, green, blues)
        dx = picPaleta.Width \ 4
        dy = picPaleta.heigt \ 4
        linha = indiceCor \ 4 + 1
        coluna = indiceCor Mod 4 + 1
        xc1 = (coluna - 1) * dx
        yc1 = (linha - 1) * dy
        xc2 = coluna * dx
        yc2 = linha * dy
        picPaleta.Line (xc1, yc1)-(xc2, yc2), , BF
        Next
End Sub

Private Sub picPaleta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case glbFerrSelec
    Case 12
        glbBcolor = picPaleta.Point(X, Y)
        frmCxFerr.picColor(0).BackColor = glbBcolor
    Case 13
        glLBcolor = picPaleta.Point(X, Y)
        frmCxFerr.picColor(1).BackColor = glbBcolor
    End Select
End Sub
