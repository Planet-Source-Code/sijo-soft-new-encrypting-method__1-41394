VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   210
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   960
      Width           =   3870
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   195
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   3900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim a
        If KeyAscii >= 65 Then
            If KeyAscii <= 90 Then
                KeyAscii = KeyAscii + 32
            End If
        End If
    a = KeyAscii
        If a = 97 Then
            a = 102
            ElseIf a = 98 Then
            a = 101
            ElseIf a = 99 Then
            a = 100
            ElseIf a = 100 Then
            a = 99
            ElseIf a = 101 Then
            a = 98
            ElseIf a = 102 Then
            a = 97
            ElseIf a = 103 Then
            a = 116
            ElseIf a = 104 Then
            a = 109
            ElseIf a = 105 Then
            a = 108
            ElseIf a = 106 Then
            a = 107
            ElseIf a = 107 Then
            a = 106
            ElseIf a = 108 Then
            a = 105
            ElseIf a = 109 Then
            a = 104
            ElseIf a = 110 Then
            a = 115
            ElseIf a = 111 Then
            a = 114
            ElseIf a = 112 Then
            a = 113
            ElseIf a = 113 Then
            a = 112
            ElseIf a = 114 Then
            a = 111
            ElseIf a = 115 Then
            a = 110
            ElseIf a = 116 Then
            a = 103
            ElseIf a = 117 Then
            a = 122
            ElseIf a = 118 Then
            a = 121
            ElseIf a = 119 Then
            a = 120
            ElseIf a = 120 Then
            a = 119
            ElseIf a = 121 Then
            a = 118
            ElseIf a = 122 Then
            a = 117
        End If
        If Not keayascii = 8 Then
            Text2.Text = Text2.Text + Chr(a)
        End If
    If KeyAscii = 8 Then
        sijo
    End If

End Sub

Private Sub sijo()
        Dim y
        y = Len(Text2.Text)
        If Not y = 0 Then
            Dim z
            z = Left(Text2.Text, y - 2)
            Text2.Text = z
        End If
End Sub
