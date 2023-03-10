VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   9720
   ClientTop       =   5865
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   10515
   Begin VB.CommandButton Command12 
      Caption         =   "SHOWALL"
      Height          =   495
      Left            =   6240
      TabIndex        =   21
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "RESET"
      Height          =   495
      Left            =   4080
      TabIndex        =   20
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "RED"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "BLUE"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "PINK"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "WHITE"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ORANGE"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "VIOLET"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "YELLOW"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00400000&
      Caption         =   "INDIGO"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8760
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GREEN"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BLACK"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7920
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label9 
      BackColor       =   &H00400000&
      Height          =   495
      Left            =   8760
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   2880
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C000&
      Height          =   495
      Left            =   3720
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0FF&
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5400
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Height          =   495
      Left            =   6240
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Visible = True
Command1.Enabled = False
Command10.Enabled = True


End Sub

Private Sub Command10_Click()
Label1.Visible = False
Label8.Visible = True
Command10.Enabled = False
Command9.Enabled = True



End Sub

Private Sub Command11_Click()
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False








End Sub

Private Sub Command12_Click()
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True

End Sub

Private Sub Command2_Click()
Label7.Visible = False
Label6.Visible = True
Command2.Enabled = False
Command8.Enabled = True




End Sub

Private Sub Command3_Click()
Label10.Visible = False
Label9.Visible = True
Command3.Enabled = False
Command4.Enabled = False



End Sub

Private Sub Command4_Click()
Label2.Visible = False
Label10.Visible = True
Command3.Enabled = True
Command5.Enabled = False




End Sub

Private Sub Command5_Click()
Label3.Visible = False
Label2.Visible = True
Command5.Enabled = False
Command4.Enabled = True



End Sub

Private Sub Command6_Click()
Label4.Visible = False
Label3.Visible = True
Command6.Enabled = False
Command5.Enabled = True



End Sub

Private Sub Command7_Click()
Label5.Visible = False
Label4.Visible = True
Command7.Enabled = False
Command6.Enabled = True




End Sub

Private Sub Command8_Click()
Label6.Visible = False
Label5.Visible = True
Command8.Enabled = False
Command7.Enabled = True




End Sub

Private Sub Command9_Click()
Label8.Visible = False
Label7.Visible = True
Command9.Enabled = False
Command2.Enabled = True



End Sub

