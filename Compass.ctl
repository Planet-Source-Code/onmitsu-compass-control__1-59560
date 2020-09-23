VERSION 5.00
Begin VB.UserControl Compass 
   BackColor       =   &H00000000&
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   ScaleHeight     =   1470
   ScaleWidth      =   1500
   Begin VB.Label NEIndex 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape Center 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   240
      Shape           =   3  'Circle
      Top             =   240
      Width           =   465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   60
      X2              =   1320
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Outer 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   480
      Shape           =   3  'Circle
      Top             =   1020
      Width           =   375
   End
End
Attribute VB_Name = "Compass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, _
ByVal bRedraw As Boolean) As Long

Dim ANGLE As Integer
Const Pi As Single = 3.1416
Const RadConv As Single = Pi / 180
Dim Radius As Integer
Dim AngleDisplay As Boolean


Sub CreateCenter()
    With Center
        .Top = (Height / 2) - (.Height / 2)
        .Left = (Width / 2) - (.Width / 2)
    End With
End Sub

Sub MoldOuter()
    Dim W As Integer
    W = Width - 120
    With Outer
        .Width = W
        .Height = W
        .Top = (Height / 2) - (.Height / 2)
        .Left = (Width / 2) - (.Width / 2)
        Radius = W / 2
    End With
End Sub

Public Sub Rotate(B As Double)
    Dim X1 As Integer
    Dim X2 As Integer
    Dim Y1 As Integer
    Dim Y2 As Integer
    
    X1 = 500
    Y1 = Outer.Top
    
    X2 = Radius * Sin(B * RadConv)
    Y2 = -1 * Radius * Cos(B * RadConv)
    
    Line1.X1 = X2 + 500
    Line1.Y1 = Y2 + 500
    
If AngleDisplay = True Then

    NEIndex.Caption = B & Chr(186)

Else

Select Case True
    Case B >= 349 Or B <= 11
        NEIndex.Caption = "N"
    Case B >= 12 And B <= 33
        NEIndex.Caption = "NNE"
    Case B >= 34 And B <= 56
        NEIndex.Caption = "NE"
    Case B >= 57 And B <= 78
        NEIndex.Caption = "ENE"
    Case B >= 79 And B <= 101
        NEIndex.Caption = "E"
    Case B >= 10 And B <= 123
        NEIndex.Caption = "ESE"
    Case B >= 124 And B <= 146
        NEIndex.Caption = "SE"
    Case B >= 147 And B <= 168
        NEIndex.Caption = "SSE"
    Case B >= 169 And B <= 191
        NEIndex.Caption = "S"
    Case B >= 192 And B <= 213
        NEIndex.Caption = "SSW"
    Case B >= 214 And B <= 236
        NEIndex.Caption = "SW"
    Case B >= 237 And B <= 258
        NEIndex.Caption = "WSW"
    Case B >= 259 And B <= 281
        NEIndex.Caption = "W"
    Case B >= 282 And B <= 303
        NEIndex.Caption = "WNW"
    Case B >= 304 And B <= 326
        NEIndex.Caption = "NW"
    Case Else
        NEIndex.Caption = "NNW"
End Select

End If

End Sub

Private Sub UserControl_Click()
If AngleDisplay = True Then
    AngleDisplay = False
Else
    AngleDisplay = True
End If
End Sub

Private Sub UserControl_Initialize()
    SetWindowRgn hwnd, _
    CreateEllipticRgn(3, 3, 65, 65), True
End Sub

Private Sub UserControl_Resize()
    If Width <> 1000 Then Width = 1000
    If Height <> 1000 Then Height = 1000

    CreateCenter
    MoldOuter

    SetLines
    
        NEIndex.Top = 400
        NEIndex.Left = 275
        NEIndex.ZOrder
End Sub

Sub SetLines()
    Line1.Y2 = 500
    Line1.X2 = 500
    If ANGLE = 0 Then
        Line1.X1 = 500
        Line1.Y1 = Outer.Top
    End If
End Sub

Public Property Get Value() As Double
    Value = Value1
End Property

Public Property Let Value(newValue As Double)
    PropertyChanged "Value"
    Rotate (newValue)
End Property


Public Property Get TextColor() As OLE_COLOR
    TextColor = NEIndex.ForeColor
End Property

Public Property Let TextColor(newTextColor As OLE_COLOR)
    NEIndex.ForeColor = newTextColor
    PropertyChanged "TextColor"
End Property

Public Property Get IndicatorColor() As OLE_COLOR
    IndicatorColor = Line1.BorderColor
End Property

Public Property Let IndicatorColor(newIndicatorColor As OLE_COLOR)
    Line1.BorderColor = newIndicatorColor
    PropertyChanged "IndicatorColor"
End Property


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        Line1.BorderColor = PropBag.ReadProperty("IndicatorColor", vbRed)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TextColor", NEIndex.ForeColor, vbWhite)
    Call PropBag.WriteProperty("IndicatorColor", Line1.BorderColor, vbRed)
End Sub
