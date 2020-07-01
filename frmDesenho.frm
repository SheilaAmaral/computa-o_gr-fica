VERSION 5.00
Begin VB.Form frmDesenho 
   Caption         =   "Desenho"
   ClientHeight    =   6360
   ClientLeft      =   525
   ClientTop       =   1980
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   596
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   726
      TabIndex        =   0
      Top             =   0
      Width           =   10950
   End
End
Attribute VB_Name = "frmDesenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CircIncrem(x1, y1, r)
n = Log(r) / Log(2) + 1
e = 1 / (2 ^ n)
xn = 0
yn = r
Picture1.PSet (x1 + Int(xn), y1 + Int(ym)), glbLColor
flag = 0
While flag < 2
    xm = xn + e * yn
    ym = -e * xm + yn
    Picture1.PSet (x1 + Int(xm), y1 + Int(ym)), glbLColor
    If (xm < e) And (ym > r - e) Then flag = flag + 1
    xn = xm
    yn = ym
Wend
End Sub
Private Sub CircTrig(x1, y1, r)
    Const pi = 3.141592654
    Dim xn As Double
    Dim yn As Double
    Dim xm As Double
    Dim ym As Double
    n = 360
    da = (2 * pi) / n
    e = Cos(da)
    f = Sin(da)
    xn = 0
    yn = r
    Picture1.PSet (x1 + Int(xn), y1 + Int(yn)), glbLColor
    For i = 1 To n
        xm = xn * e - yn * f
        ym = xn * f + yn * e
        Picture1.PSet (x1 + Int(xm), y1 + Int(ym)), glbLColor
        xn = xm
        yn = ym
        Next i
End Sub
Private Sub poliregular(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
pi = 3.141592
n = glbFiltro
alfa = 2 * pi / n
dx = x2 - x1
dy = y2 - y1
r = Sqr(dx * dx + dy * dy)
l = 2 * r * Sin(alfa / 2)
beta = alfa
xi = x1 + r * Sin(alfa / 2)
yi = y1 - r * Cos(alfa / 2)
For i = 1 To n
    xf = xi + l * Cos(beta)
    yf = yi + l * Sin(beta)
    Picture1.Line (yi, yi)-(xf, yf), glbLColor
    xi = xf
    yi = yf
    beta = beta + alfa
    Next i
End Sub
Private Sub Bresenham(x1 As Integer, y1asinteger, x2 As Integer, y2 As Integer)
    dx = Abs(x2 - x1)
    dy = Abs(y2 - y1)
    mov = 0
    
    If dx <> 0 Then e = (dy / dx) - 0.5
    da = dx
    If dy > dx Then
        mov = 1
        e = (dx / dy) - 0.5
        da = dy
    End If
    
    X = x1
    Y = y1
    incx = 1
    If x1 > x2 Then incx = -1
    If y1 > y2 Then incy = -1
    
    For i = 1 To da
        Picture1.PSet (X, Y)
        If e > 0 Then
            If mov = 0 Then
                Y = Y + incy
            Else
                X = X + incx
            End If
        
            e = e - 1
        End If
        If mov = 0 Then
            X = X + incx
            e = e + dy / dx
        Else
            Y = Y + incy
            e = e + dx / dy
        End If
   nexti
    
End Sub

Private Sub DDA(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)

dx = Abs(x2 - x1)
dy = Abs(y2 - y1)
If dy > dx Then dx = dy
If dx <> 0 Then
    Ax = (x2 - x1) / dx
    Ay = (y2 - y1) / dx
    X = x1 + 0.5
    Y = y1 + 0.5
    For i = 1 To dx
        Picture1.PSet (X, Y)
        X = X + Ax
        Y = Y + Ay
    Next
End If

End Sub
Private Sub Form_Load()

Me.Left = 1350
Me.Top = 0
Me.Width = 11070
Me.Height = 6940
glbLColor = QBColor(5)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    x1 = X
    y1 = Y
    x2 = X
    y2 = Y
    'glbTcolor = Picture1.Point(x1, y1)
    Picture1.DrawMode = 13
    If glbFerrSelec = 3 Then Picture1.PSet (x1, y1)
    glbProxPonto = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Tcolor As Long
    Dim Vcolor As Long
    x3 = x2
    y3 = y2
    x2 = X
    y2 = Y
    Picture1.AutoRedraw = False
    If glbFerrSelec <> 3 And glbFerrSelec <> 6 Then Picture1.DrawMode = 7
    'Vcolor = glbTcolor
    'glbTcolor = Picture1.Point(x2, y2)
    'Tcolor = &HFFFFFF - glbTcolor
    If glbProxPonto = True Then
        Select Case glbFerrSelec
            Case 0
                Picture1.Line (x1, y1)-(x2, y2), glbLColor
                Picture1.Line (x1, y1)-(x3, y3), glbLColor
            Case 1 To 2
                Picture1.Line (x1, y1)-(x2, y2), glbLColor
                Picture1.Line (x1, y1)-(x3, y3), glbLColor
            Case 3
                Picture1.AutoRedraw = True
                Picture1.Line -(x3, y3), glbLColor
                Picture1.Line -(x2, y2), glbLColor
            Case 4
                Select Case glbMetodo
            Case 11
                Picture1.Line (x1, y1)-(x2, y2), glbLColor, B
                Picture1.Line (x1, y1)-(x3, y3), glbLColor, B
            Case Else
                dx = x2 - x1
                dy = y2 - y1
                r = Sqr(dx * dx + dy * dy)
                Picture1.Circle (x1, y1), r, glbLColor
                dx = x3 - x1
                dy = y3 - y1
                r = Sqr(dx * dx + dy * dy)
                Picture1.Circle (x1, y1), r, glbLColor
            End Select
            Case 5
                xm = (x1 + x2) / 2
                ym = (y1 + y2) / 2
                s1 = Abs(x2 - x1) / 2
                s2 = Abs(y2 - y1) / 2
                If s1 < s2 Then
                    r = s1
                Else
                    r = s2
                End If
                If s1 = 0 Then
                    a = 1
                    If s2 = 0 Then
                        Exit Sub
                    End If
                Else
                    a = s2 / s1
                End If
                Picture1.Circle (xm, ym), r, glbLColor, , , a
                xm = (x1 + x3) / 2
                ym = (y1 + y3) / 2
                s1 = Abs(x3 - x1) / 2
                s2 = Abs(y3 - y1) / 2
                If s1 < s2 Then
                    r = s1
                Else
                    r = s2
                End If
                If s1 = 0 Then
                    a = 1
                    If s2 = 0 Then
                        'R=1
                        Exit Sub
                    End If
                Else
                    a = s2 / s1
                End If
                Picture1.Circle (xm, ym), r, glbLColor, , , a
            Case 6
                Picture1.AutoRedraw = True
                glbBcolor = RGB(255, 255, 255)
                Picture1.Line (x2 - 10, y2 - 10)-(x2 + 10, y2 + 10), glbBcolor, BF
            Case Else
                Picture1.Line (x1, y1)-(x2, y2), glbLColor, B
                Picture1.Line (x1, y1)-(x3, y3), glbLColor, B
            End Select
        End If
 End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    x2 = X
    y2 = Y
    glbProxPonto = False
    Picture1.DrawMode = 13
    Picture1.AutoRedraw = True
    Select Case glbFerrSelec
    Case 0 'LinhaReta
        Select Case glbMetodo
        Case 0 'DDA
            DDA x1, y1, x2, y2
        Case 1 'Bresenham
            Bresenham x1, y1, x2, y2
        Case 2 'Metodo VB
            Picture1.Line (x1, y1)-(x2, y2), glbLColor
        End Select
    Case 1 'Poligonos vazados
        Select Case glbMetodo
        Case 3 'Triangulo
            xa = x1
            ya = y2
            d = x2 - x1
            xb = x1 + d / 2
            yb = y1
            xc = x2
            yc = y2
            Picture1.Line (xa, ya)-(xb, yb), glbLColor
            Picture1.Line (xb, yb)-(xc, yc), glbLColor
            Picture1.Line (xc, yc)-(xa, ya), glbLColor
        Case 4 'Retangulo
            Picture1.Line (x1, y1)-(x2, y2), glbLColor, B
        Case 5 'Pentagono
            d = x2 - x1
            l = y2 - y1
            xa = x1 + d / 2
            ya = y1
            xb = x2
            yb = y2 - l / 2 - l / 8
            xc = x2 - d / 4
            yc = y2
            xd = x1 + d / 4
            yd = y2
            xe = x1
            ye = yb
            Picture1.Line (xa, ya)-(xb, yb), glbLColor
            Picture1.Line (xb, yb)-(xc, yc), glbLColor
            Picture1.Line (xc, yc)-(xd, yd), glbLColor
            Picture1.Line (xd, yd)-(xe, ye), glbLColor
            Picture1.Line (xe, ye)-(xa, ya), glbLColor
        Case 6 'Estrela
            dx = x2 - x1
            dy = y2 - y1
            xa = x1 + dx / 2
            ya = y1
            xb = x1 + 5 * dx / 8
            yb = y1 + 3 * dy / 8
            xc = x2
            yc = yb
            xd = x1 + 11 * dx / 16
            yd = y1 + 5 * dy / 8
            xe = x1 + 13 * dx / 16
            ye = x2
            xf = xa
            yf = y1 + 6 * dy / 8
            xg = x1 + 3 * dx / 16
            yg = y2
            xh = x1 + 5 * dx / 16
            yh = yd
            xi = x1
            yi = yb
            xj = x1 + 3 * dx / 8
            yj = yb
            Picture1.Line (xa, ya)-(xb, yb), glbLColor
            Picture1.Line (xb, yb)-(xc, yc), glbLColor
            Picture1.Line (xc, yc)-(xd, yd), glbLColor
            Picture1.Line (xd, yd)-(xe, ye), glbLColor
            Picture1.Line (xe, ye)-(xf, yf), glbLColor
            Picture1.Line (xf, yf)-(xg, yg), glbLColor
            Picture1.Line (xg, yg)-(xh, yh), glbLColor
            Picture1.Line (xh, yh)-(xi, yi), glbLColor
            Picture1.Line (xi, yi)-(xj, yj), glbLColor
            Picture1.Line (xj, yj)-(xa, ya), glbLColor
         Case 7 'regulares
            poliregular x1, y1, x2, y2
        End Select
    Case 2 'poligonos cheios
        Picture1.Line (x1, y1)-(x2, y2), glbLColor, BF
    Case 3 'traço livre
        Picture1.Line -(x2, y2), glbLColor
    Case 4 'círculos
        dx = x2 - x1
        dy = y2 - y1
        r = Sqr(dx * dx + dy * dy)
        Select Case glbMetodo
        Case 8
            CircTrig x1, y1, r
        Case 9
            CircIncrem x1, y1, r
        Case 10
            Picture1.Circle (x1, y1), r, glbLColor
    Case 4 'elipses
        xm = (x1 + x2) / 2
        ym = (y1 + y2) / 2
        s1 = Abs(x2 - x1) / 2
        s2 = Abs(y2 - y1) / 2
        If s1 < s2 Then
            r = s2
        Else
            r = s1
        End If
        If s1 = 0 Then
            a = 1
            Exit Sub
        Else
            a = s2 / s1
        End If
        Picture1.Circle (xm, ym), r, glbLColor, , , a
    Case 5 'arcos
        a = -dx / dy
        Dim dx2 As Double
        Dim dx1 As Double
        Dim dy2 As Double
        Dim dy1 As Double
        dx2 = x2
        dx2 = dx2 * dx2
        dx1 = x1
        dx1 = dx1 * x1
        dx2 = dx2 - dx1
        dy2 = y2
        dy2 = dy2 * y2
        dy1 = y1
        dy1 = dy1 * y1
        dy2 = dy2 - dy1
        B = (dx2 + dy2) / (2 * dy)
        If Abs(dx) > Abs(dy) Then
            xc = x1
            yc = a * xc + B
            a2 = 4 * Atn(1) / 2
            a1 = Atn((yc - y2) / (x2 - xc))
            r = Abs(yc - y1)
            If dx < 0 Then
                If dy > 0 Then
                    a2 = 4 * Atn(1) + a1
                    a1 = 4 * Atn(1) / 2
                Else
                    a1 = 4 * Atn(1) + a1
                    a2 = 6 * Atn(1)
                End If
            End If
            If dy < 0 And dx > 0 Then
                a2 = 8 * Atn(1) + a1
                a1 = 6 * Atn(1)
            End If
            Else
                yc = y2
                xc = (yc - B) / a
                a1 = 0
                a2 = Atn((yc - y1) / (x1 - xc))
                r = Abs(x2 - xc)
                If dy < 0 Then
                    If dx > 0 Then
                        a1 = 8 * Atn(1) + a2
                        a2 = 8 * Atn(1)
                    Else
                        a1 = 4 * Atn(1)
                        a2 = 6 * Atn(1) - a2
                    End If
                End If
                If dx < 0 And dy > 0 Then
                    a1 = 4 * Atn(1) + a2
                    a2 = 4 * Atn(1)
                End If
            End If
            Picture1.Circle (xc, yc), r, glbLColor, a1, a2
         End Select
    'DDA x1, y1, x2, y2
    End Select
End Sub
