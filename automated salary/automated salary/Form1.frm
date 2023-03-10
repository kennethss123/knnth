VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   11175
   ClientTop       =   3990
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   5910
   Begin VB.CommandButton Command2 
      Caption         =   "MANAGER"
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
      Left            =   240
      TabIndex        =   27
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "POSITION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   26
      Top             =   5760
      Width           =   5655
      Begin VB.CommandButton Command7 
         Caption         =   "JANITOR"
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
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "SUPER VISOR"
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
         Left            =   2160
         TabIndex        =   31
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "HELPER"
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
         Left            =   2160
         TabIndex        =   30
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "MASON"
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
         Left            =   4320
         TabIndex        =   29
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "STREET SWEEPER"
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
         Left            =   4320
         TabIndex        =   28
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   25
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   3720
      TabIndex        =   24
      Text            =   "0"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   3720
      TabIndex        =   23
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   1080
      TabIndex        =   22
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   1080
      TabIndex        =   21
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   1080
      TabIndex        =   20
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NET PAY"
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
      Left            =   2400
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SALARY INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5655
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   3600
         TabIndex        =   19
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   3600
         TabIndex        =   18
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   3600
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   960
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   960
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "GROSS PAY"
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OVERTIME"
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "HOLIDAY"
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO. OF   DAYS"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RATE"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "POSITION"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DEDUCTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   5655
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CASH ADVANCE"
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PHILHEALTH"
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PAGIBIG"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SSS"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TAX 6.2%"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label13_Click()
End Sub

Private Sub Command1_Click()
Text6.Text = Val(Text2.Text) * Val(Text3.Text) + (Val(Text4.Text) * (Val(Text2.Text) - 90) + (Val(Text5.Text) * ((Val(Text2.Text) / 8) + (Val(Text2.Text) / 8) * 0.3)))
Text7.Text = Val(Text6.Text) * 0.062
Text8.Text = Val(Text6.Text) * 0.06
Text9.Text = 100
Text10.Text = 100
Text12.Text = Val(Text6.Text) - (Val(Text7.Text) + (Val(Text8.Text) + (Val(Text9.Text) + (Val(Text10.Text) + Val(Text11.Text)))))
End Sub

Private Sub Command2_Click()
Text2.Text = 1500
Text1.Text = "MANAGER"
End Sub

Private Sub Command3_Click()
Text2.Text = 550
Text1.Text = "STREET SWEEPER"
End Sub

Private Sub Command4_Click()
Text2.Text = 600
Text1.Text = "MASON"
End Sub

Private Sub Command5_Click()
Text2.Text = 500
Text1.Text = "HELPER"
End Sub

Private Sub Command6_Click()
Text2.Text = 2000
Text1.Text = "SUPER VISOR"
End Sub

Private Sub Command7_Click()
Text2.Text = 500
Text1.Text = "JANITOR"
End Sub
