VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5190
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5190
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   15
      Left            =   2880
      TabIndex        =   15
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   14
      Left            =   2040
      TabIndex        =   14
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   13
      Left            =   1200
      TabIndex        =   13
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   12
      Left            =   360
      TabIndex        =   12
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   11
      Left            =   2880
      TabIndex        =   11
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   10
      Left            =   2040
      TabIndex        =   10
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   9
      Left            =   1200
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   8
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   7
      Left            =   2880
      TabIndex        =   7
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   6
      Left            =   2040
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   3
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n(1 To 4, 1 To 4) As Single
Sub output()
For i = 1 To 4
    For j = 1 To 4
        If n(i, j) = 0 Then
            Label1((i - 1) * 4 + j - 1).Caption = ""
        Else
            Label1((i - 1) * 4 + j - 1).Caption = n(i, j)
            Select Case n(i, j)
                Case 2
                    Label1((i - 1) * 4 + j - 1).ForeColor = &HC0C0C0
                Case 4
                    Label1((i - 1) * 4 + j - 1).ForeColor = &H808080
                Case 8
                    Label1((i - 1) * 4 + j - 1).ForeColor = &H404040
                Case 16
                    Label1((i - 1) * 4 + j - 1).ForeColor = &H0&
                Case 32
                    Label1((i - 1) * 4 + j - 1).ForeColor = &H400040
                Case 64
                    Label1((i - 1) * 4 + j - 1).ForeColor = &H800080
                Case 128
                    Label1((i - 1) * 4 + j - 1).ForeColor = &HC000C0
                Case 256
                    Label1((i - 1) * 4 + j - 1).ForeColor = &HFF00FF
                Case 512
                    Label1((i - 1) * 4 + j - 1).ForeColor = &HFF80FF
                Case 1024
                    Label1((i - 1) * 4 + j - 1).ForeColor = &HFF8080
                Case 2048
                    Label1((i - 1) * 4 + j - 1).ForeColor = &HFF0000
            End Select
        End If
    Next j
Next i

End Sub
Sub check_rd()
flag = 1
For i = 1 To 4
    For j = 1 To 4
        If n(i, j) = 0 Then flag = 0
    Next j
Next i
If flag = 0 Then rd
End Sub

Sub rd()
Randomize
x = Int(Rnd * (4 - 1 + 1)) + 1
y = Int(Rnd * (4 - 1 + 1)) + 1
If n(x, y) = 0 Then
    n(x, y) = 2
Else
    rd
End If
output
End Sub
Sub check()
flag = 0
For i = 1 To 4
    For j = 1 To 4
        If n(i, j) = 0 Then flag = 1
    Next j
Next i
If flag = 0 Then
    For i = 1 To 4
        For j = 1 To 3
            If n(i, j) = n(i, j + 1) Then flag = 1
        Next j
    Next i
    For i = 1 To 3
        For j = 1 To 4
            If n(i, j) = n(i + 1, j) Then flag = 1
        Next j
    Next i
End If
If flag = 0 Then MsgBox ("Better luck next time!")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
check
flag = 0
Select Case KeyCode
    Case vbKeyDown
        For i = 4 To 2 Step -1
            For j = 1 To 4
                k = i - 1
                Do While (k > 1) And (n(k, j) = 0)
                    k = k - 1
                Loop
                If n(i, j) = n(k, j) Then
                    n(i, j) = 2 * n(i, j)
                    n(k, j) = 0
                    flag = 1
                End If
            Next j
        Next i
        For i = 3 To 1 Step -1
            For j = 1 To 4
                If n(i + 1, j) = 0 Then
                    k = i
                    Do While (k > 1) And (n(k, j) = 0)
                        k = k - 1
                    Loop
                    n(i + 1, j) = n(k, j)
                    n(k, j) = 0
                    flag = 1
                End If
            Next j
        Next i
    Case vbKeyUp
        For i = 1 To 3
            For j = 1 To 4
                k = i + 1
                Do While (k < 4) And (n(k, j) = 0)
                    k = k + 1
                Loop
                If n(i, j) = n(k, j) Then
                    n(i, j) = 2 * n(i, j)
                    n(k, j) = 0
                    flag = 1
                End If
            Next j
        Next i
        For i = 2 To 4
            For j = 1 To 4
                If n(i - 1, j) = 0 Then
                    k = i
                    Do While (k < 4) And (n(k, j) = 0)
                        k = k + 1
                    Loop
                    n(i - 1, j) = n(k, j)
                    n(k, j) = 0
                    flag = 1
                End If
            Next j
        Next i
    Case vbKeyLeft
        For i = 1 To 4
            For j = 1 To 3
                k = j + 1
                Do While (k < 4) And (n(i, k) = 0)
                    k = k + 1
                Loop
                If n(i, j) = n(i, k) Then
                    n(i, j) = 2 * n(i, j)
                    n(i, k) = 0
                    flag = 1
                End If
            Next j
        Next i
        For i = 1 To 4
            For j = 2 To 4
                If n(i, j - 1) = 0 Then
                    k = j
                    Do While (k < 4) And (n(i, k) = 0)
                        k = k + 1
                    Loop
                    n(i, j - 1) = n(i, k)
                    n(i, k) = 0
                    flag = 1
                End If
            Next j
        Next i
    Case vbKeyRight
        For i = 1 To 4
            For j = 4 To 2 Step -1
                k = j - 1
                Do While (k > 1) And (n(i, k) = 0)
                    k = k - 1
                Loop
                If n(i, j) = n(i, k) Then
                    n(i, j) = 2 * n(i, j)
                    n(i, k) = 0
                    flag = 1
                End If
            Next j
        Next i
        For i = 1 To 4
            For j = 3 To 1 Step -1
                If n(i, j + 1) = 0 Then
                    k = j
                    Do While (k > 1) And (n(i, k) = 0)
                        k = k - 1
                    Loop
                    n(i, j + 1) = n(i, k)
                    n(i, k) = 0
                    flag = 1
                End If
            Next j
        Next i
    End Select
If flag = 1 Then check_rd
output
End Sub

Private Sub Form_Load()
rd
rd
output
End Sub
