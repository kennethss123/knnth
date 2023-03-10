VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H008080FF&
   ClientHeight    =   7905
   ClientLeft      =   5760
   ClientTop       =   3990
   ClientWidth     =   17160
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   17160
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   42
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   41
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   40
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   39
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   38
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   37
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   36
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   35
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   33
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "DINE OUT"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      TabIndex        =   32
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "DINE  IN"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      TabIndex        =   31
      Top             =   4800
      Width           =   1215
   End
   Begin VB.OptionButton Option5 
      Caption         =   "THICK CRUST"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   30
      Top             =   3360
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "THIN CRUST"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   29
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   6840
      TabIndex        =   24
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   6840
      TabIndex        =   23
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   11280
      TabIndex        =   22
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   11280
      TabIndex        =   21
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   11280
      TabIndex        =   20
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   6840
      TabIndex        =   19
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H0080FFFF&
      Caption         =   "EXTRA CHEESE"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   18
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H0080FFFF&
      Caption         =   "TOMATOES"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   17
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H0080FFFF&
      Caption         =   "MUSHROOMS"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   16
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "exit"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7560
      TabIndex        =   11
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "reset"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      TabIndex        =   10
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      Caption         =   "TOPPINGS"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4920
      TabIndex        =   9
      Top             =   4320
      Width           =   7695
      Begin VB.CheckBox Check6 
         BackColor       =   &H0080FFFF&
         Caption         =   "ONIONS"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   27
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H0080FFFF&
         Caption         =   "GREEN PEPPERS"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   26
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H0080FFFF&
         Caption         =   "BLACK OLIVES"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   25
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF80FF&
      Caption         =   "SIZE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   8880
      TabIndex        =   7
      Top             =   1920
      Width           =   3735
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF0000&
         Caption         =   "CRUST TYPES"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   3735
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   22.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   1560
            TabIndex        =   15
            Top             =   840
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF80FF&
      Caption         =   "SIZE"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4920
      TabIndex        =   6
      Top             =   1920
      Width           =   3735
      Begin VB.OptionButton Option1 
         Caption         =   "SMALL"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "LARGE"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "MEDUIM"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   5
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "create pizza"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      MaskColor       =   &H008080FF&
      TabIndex        =   2
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "TOPPINGS"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   45
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CRUST TYPE"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   44
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SIZE"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   43
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PIZZA ORDER "
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   5280
      TabIndex        =   1
      Top             =   480
      Width           =   7815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   1455
      Left            =   4920
      TabIndex        =   0
      Top             =   360
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = vbChecked Then
 Check2.Value = vbUnchecked
 Text18.Text = "DINE IN"
 Text18.Visible = True
 End If
 If Check1.Value = vbUnchecked Then
 Text18.Text = ""
 Text18.Visible = False
 End If
 
 
End Sub

Private Sub Check2_Click()
If Check2.Value = vbChecked Then
 Check1.Value = vbUnchecked
 Text18.Text = "DINE OUT"
 Text18.Visible = True
 End If
 If Check2.Value = vbUnchecked Then
 Text18.Text = ""
 Text18.Visible = False
 End If
End Sub

Private Sub Check3_Click()
If Check3.Value = vbChecked Then
Text4.Text = "30"
Text14.Text = "MUSHROOMS"
Text14.Visible = True

End If

If Check3.Value = vbUnchecked Then
Text4.Text = ""
Text14.Text = ""
Text14.Visible = False

End If


End Sub

Private Sub Check4_Click()
If Check4.Value = vbChecked Then
Text7.Text = "50"
Text17.Text = "BLACK OLIVES"
Text17.Visible = True
End If
If Check4.Value = vbUnchecked Then
Text7.Text = ""
Text17.Text = "TOMATOES"
Text17.Visible = False
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = vbChecked Then
Text6.Text = "40"
Text15.Text = "GREEN PEPPERS"
Text15.Visible = True
End If
If Check5.Value = vbUnchecked Then
Text6.Text = ""
Text15.Text = ""
Text15.Visible = False
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = vbChecked Then
Text5.Text = "20"
Text13.Text = "ONIONS"
Text13.Visible = True
End If
If Check6.Value = vbUnchecked Then
Text5.Text = ""
Text13.Text = ""
Text13.Visible = False

End If
End Sub

Private Sub Check8_Click()
If Check8.Value = vbChecked Then
Text8.Text = "40"
Text21.Text = "TOMATOES"
Text21.Visible = True
End If
If Check8.Value = vbUnchecked Then
Text8.Text = ""
Text21.Text = ""
Text21.Visible = False
End If
End Sub
Private Sub Check9_Click()
If Check9.Value = vbChecked Then
Text9.Text = "30"
Text12.Text = "EXTRA CHEESE"
Text12.Visible = True
End If
If Check9.Value = vbUnchecked Then
Text9.Text = ""
Text12.Text = ""
Text12.Visible = False

End If
End Sub

Private Sub Command1_Click()
Dim message As String
Text1.Text = Val(Text2) + Val(Text3) + Val(Text4) + Val(Text5) + Val(Text6) + Val(Text7) + Val(Text8) + Val(Text9) + 20

MsgBox "TOTAL AMOUNT=" + Text1.Text +



End Sub

Private Sub Command2_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Text2.Text = ""
Text3.Text = ""
Text1.Text = ""
Text12.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text20.Text = ""

Check1 = vbUnchecked
Check2 = vbUnchecked
Check3 = vbUnchecked
Check4 = vbUnchecked
Check5 = vbUnchecked
Check6 = vbUnchecked
Check7 = vbUnchecked
Check8 = vbUnchecked
Check9 = vbUnchecked




End Sub

Private Sub Command3_Click()
Me.Hide


End Sub

Private Sub Command4_Click()
Text1.Text = Val(Text2) + Val(Text3) + Val(Text4) + Val(Text5) + Val(Text6) + Val(Text7) + Val(Text8) + Val(Text9) + 20

End Sub

Private Sub Command5_Click()

End If



End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Text2.Text = "20"
Text20.Text = "SMALL"
Text20.Visible = True

End If

End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Text2.Text = "40"
Text20.Text = "MEDIUM"
Text20.Visible = True

End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Text2.Text = "60"
Text20.Text = "LARGE"
Text20.Visible = True
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
Text3.Text = "30"
Text16.Text = "THIN CRUST"
Text16.Visible = True
End If
End Sub

Private Sub Option5_Click()
If Option5.Value = True Then
Text3.Text = "50"
Text16.Text = "THICK CRUST"
Text16.Visible = True
End If
End Sub

