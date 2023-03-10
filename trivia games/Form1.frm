VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   14595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   33
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   32
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   26
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   25
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command24 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      TabIndex        =   24
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command23 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12120
      TabIndex        =   23
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command22 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      TabIndex        =   22
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command21 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9720
      TabIndex        =   21
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command19 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   19
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   18
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   17
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   16
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   15
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   14
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   13
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      TabIndex        =   12
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12120
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9720
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "COPYRIGHT 2021"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      TabIndex        =   36
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00004040&
      BackStyle       =   0  'Transparent
      Caption         =   "TRIVIA "
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2640
      TabIndex        =   35
      Top             =   600
      Width           =   8895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DEVELOPED BY:K3N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      TabIndex        =   34
      Top             =   8760
      Width           =   3135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      TabIndex        =   31
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   30
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   29
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   28
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   27
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   14295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Caption = "A"
Command1.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Val(Text1.Text) = 5 Then
MsgBox "GOOD JOB :)"
Form2.Show
Me.Hide

End If


End Sub

Private Sub Command10_Click()
Command10.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command11_Click()
Command11.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command12_Click()
Command12.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command13_Click()
Command13.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command14_Click()
Command14.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command15_Click()
Command15.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command16_Click()
Command16.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command17_Click()
Command17.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command18_Click()
Command18.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command19_Click()
Label2.Caption = "N"
Command19.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Val(Text1.Text) = 5 Then
MsgBox "GOOD JOB :)"
Form2.Show
Me.Hide

End If
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If

End Sub

Private Sub Command20_Click()
Label4.Caption = "M"
Command20.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Val(Text1.Text) = 5 Then
MsgBox "GOOD JOB :)"
Form2.Show
Me.Hide

End If
End Sub

Private Sub Command21_Click()
Command21.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command22_Click()
Command22.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command23_Click()
Command23.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command24_Click()
Command24.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command25_Click()
Command25.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command26_Click()
Command26.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command3_Click()
Command3.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command4_Click()
Command4.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command5_Click()
Label5.Caption = "E"
Command5.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Val(Text1.Text) = 5 Then
MsgBox "GOOD JOB :)"
Form2.Show
Me.Hide

End If
End Sub

Private Sub Command6_Click()
Command6.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command7_Click()
Command7.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command8_Click()
Command8.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Val(Text2.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command9_Click()
Label3.Caption = "I"
Command9.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Val(Text1.Text) = 5 Then
MsgBox "GOOD JOB :)"
Form2.Show
Me.Hide

End If
End Sub

