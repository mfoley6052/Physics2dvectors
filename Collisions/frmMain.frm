VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Elastic Collision Sim"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtVelocity 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtmass 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtmass 
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   2640
      Width           =   735
   End
   Begin VB.Timer tmrIsColliding 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   2520
   End
   Begin VB.Label Label3 
      Caption         =   "Velocity 1"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "mass 2"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "mass 1"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2400
      Width           =   615
   End
   Begin VB.Shape shpBall 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   375
   End
   Begin VB.Shape shpBall 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   480
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type obj
    v1 As Double
    v2 As Double
End Type
Private Type obj2D
    vertical As obj
    horizontal As obj
    x As Integer
    y As Integer
    startX As Integer
    startY As Integer
    mass As Double
    height As Integer
    width As Integer
End Type
Private Type Objs1Var
    i(1) As Double
End Type
Dim ball(1) As obj2D
Dim rst As Boolean
Private Function xCollidesYAtRest(ByRef x As obj2D, ByRef y As obj2D) As Objs1Var
x.horizontal.v2 = x.horizontal.v1 * ((x.mass - y.mass) / (x.mass + y.mass))
y.horizontal.v2 = x.horizontal.v1 * ((2 * x.mass) / (x.mass + y.mass))
xCollidesYAtRest.i(0) = x.horizontal.v2
xCollidesYAtRest.i(1) = y.horizontal.v2
End Function
    
Private Sub cmdGo_Click()
If IsNumeric(txtmass(0).Text) And IsNumeric(txtmass(1).Text) And IsNumeric(txtVelocity.Text) Then
    If rst Then
        For q = shpBall.LBound To shpBall.UBound
            ball(q).mass = txtmass(q).Text
        Next q
        ball(0).horizontal.v1 = txtVelocity.Text
        rst = False
        cmdGo.Caption = "Reset"
        tmrIsColliding.Enabled = True
    Else
        For q = shpBall.LBound To shpBall.UBound
            ball(q).x = ball(q).startX
            ball(q).y = ball(q).startY
            shpBall(q).Left = ball(q).x - shpBall(q).width
            'shpBall(q).Top = ball(q).y - shpBall(q).height
        Next q
        ball(1).horizontal.v1 = 0
        ball(1).horizontal.v2 = 0
        
        cmdGo.Caption = "Go"
        rst = True
        tmrIsColliding.Enabled = False
    End If
End If
End Sub

Private Sub Form_Load()
For x = shpBall.LBound To shpBall.UBound
    ball(x).mass = shpBall(x).width * 3.14
    ball(x).x = shpBall(x).Left + shpBall(x).width
    ball(x).y = shpBall(x).Top + shpBall(x).height
    ball(x).startX = ball(x).x
    ball(x).startY = ball(x).y
    ball(x).width = shpBall(x).width
    ball(x).height = shpBall(x).height
Next x
rst = True
ball(0).horizontal.v1 = 10
End Sub

Private Sub GameUpdate()
For x = shpBall.LBound To shpBall.UBound
    shpBall(x).Left = ball(x).x
    'shpBall(x).Top = ball(x).y
Next x
End Sub
Private Function obj2objConv(ByRef temp As Objs1Var, ByVal first As Boolean) As Double
Dim tempo(1) As String
For x = shpBall.LBound To shpBall.UBound
tempo(x) = Str(temp.i(x))
Next x
If first Then
obj2objConv = Val(tempo(0))
Else
obj2objConv = Val(tempo(1))
End If
End Function


Private Sub tmrIsColliding_Timer()
Dim temp As Objs1Var
For q = shpBall.LBound To shpBall.UBound
    ball(q).x = ball(q).x + ball(q).horizontal.v1
    'ball(q).y = ball(q).y + ball(q).vertical.v1
Next q
If ball(0).x >= ball(1).x - ball(1).width And ball(0).x <= ball(1).x + ball(1).width Then
    temp = xCollidesYAtRest(ball(0), ball(1))
    ball(0).horizontal.v2 = obj2objConv(temp, True)
    ball(1).horizontal.v2 = obj2objConv(temp, False)
    For q = shpBall.LBound To shpBall.UBound
        ball(q).horizontal.v1 = ball(q).horizontal.v2
    Next q
End If
Call GameUpdate
End Sub
