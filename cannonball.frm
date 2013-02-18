VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.Line lnPointer 
      X1              =   120
      X2              =   720
      Y1              =   4680
      Y2              =   4200
   End
   Begin VB.Shape shpBall 
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Vertical
    d As Single
    V1 As Single
    V2 As Single
    a As Single
    t As Single
End Type

Private Type Horizontal
    d As Single
    V1 As Single
    V2 As Single
    a As Single
    t As Single
End Type

Dim vert As Vertical
Dim hor As Horizontal
Dim startX As Single
Dim startY As Single


Private Sub Form_Load()
lnPointer.X1 = (shpBall.Left + shpBall.Width / 2)
lnPointer.Y1 = (shpBall.Top + shpBall.Height / 2)
vert.a = 9.8
hor.a = 0
vert.V1 = 0
hor.V1 = 0
startX = shpBall.Left
startY = shpBall.Top
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpBall.Top = startY
shpBall.Left = startX
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lnPointer.X2 = X
lnPointer.Y2 = Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
vert.d = Y
hor.d = X
vert.V2 = Sqr((vert.V1 * vert.V1) + 2 * vert.a * vert.d)
vert.t = (vert.V2 - vert.V1) / vert.a
Do Until c >= vert.t
    shpBall.Top = shpBall.Top - vert.V1
    vert.V1 = vert.V1 + vert.a
    c = c + 1
Loop
hor.V2 = hor.d / vert.t
Do Until b >= vert.t
    b = b + 1
    shpBall.Left = shpBall.Left + hor.V2
Loop
End Sub
