VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Compass"
   ClientHeight    =   1200
   ClientLeft      =   5730
   ClientTop       =   5700
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   3000
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "0"
      Top             =   360
      Width           =   615
   End
   Begin Project1.Compass Compass 
      Height          =   1005
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1005
      _extentx        =   1773
      _extenty        =   1773
   End
   Begin VB.Label Label1 
      Caption         =   "Degrees"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'simple compass control that i created for a gps application im working on. 
'degrees or laymans direction can be displayed and changed during runtime, by clicking on control. 
'some source (rounded form, rotation) was grabbed from several sources on the web. 

Private Sub Text1_Change()

    If Val(Text1.Text) > 359 Then
        Text1.Text = (Val(Text1.Text) - 360)
    End If
    
    Compass.Value = Text1.Text
    

End Sub
