VERSION 5.00
Begin VB.Form frmPizza 
   Caption         =   "PIZZA ORDER"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   675
      Left            =   7200
      TabIndex        =   27
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   18
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      TabIndex        =   17
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton CmdBuild 
      Caption         =   "Build "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   16
      Top             =   5400
      Width           =   1935
   End
   Begin VB.OptionButton optWhere 
      Caption         =   "Eat Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   4440
      TabIndex        =   15
      Top             =   5280
      Width           =   1455
   End
   Begin VB.OptionButton optWhere 
      Caption         =   "Eat In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   4440
      TabIndex        =   14
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Frame Sizes 
      Caption         =   "Sizes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   3255
      Begin VB.OptionButton optSize 
         Caption         =   "Large"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton optSize 
         Caption         =   "Meduim"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optSize 
         Caption         =   "Small"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   735
         Left            =   1080
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Crust type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   960
      TabIndex        =   1
      Top             =   3600
      Width           =   3255
      Begin VB.OptionButton optCrust 
         Caption         =   "Thick Crust"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton optCrust 
         Caption         =   "Thin Crust"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H008080FF&
         Height          =   735
         Left            =   1560
         TabIndex        =   26
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Toppings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4440
      TabIndex        =   0
      Top             =   720
      Width           =   5775
      Begin VB.CheckBox chkTop 
         Caption         =   "Onions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   2640
         TabIndex        =   13
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Pork"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   2640
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Sausage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   2640
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Peperoni"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Tomatoes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Extra Cheese"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   25
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   24
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   22
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   21
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPizza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PizzaSize As String
Dim PizzaCrust As String
Dim PizzaWhere As String



Private Sub Check1_Click()

End Sub


Private Sub Command1_Click()

End Sub

Private Sub chkTop_Click(Index As Integer)
If chkTop(0) = vbChecked Then
Label4.Caption = 30
End If
If chkTop(0) = vbUnchecked Then
Label4.Caption = ""
End If
If chkTop(1) = vbChecked Then
Label5.Caption = 25
End If
If chkTop(1) = vbUnchecked Then
Label5.Caption = ""
End If
If chkTop(2) = vbChecked Then
Label6.Caption = 40
End If
If chkTop(2) = vbUnchecked Then
Label6.Caption = ""
End If
If chkTop(3) = vbChecked Then
Label7.Caption = 35
End If
If chkTop(3) = vbUnchecked Then
Label7.Caption = ""
If chkTop(4) = vbChecked Then
Label8.Caption = 30
End If
If chkTop(4) = vbUnchecked Then
Label8.Caption = ""
End If
If chkTop(5) = vbChecked Then
Label9.Caption = 25
End If
If chkTop(5) = vbUnchecked Then
Label9.Caption = ""
End If

End If
End Sub

Private Sub CmdBuild_Click()
Dim message As String
Dim I As Integer
message = PizzaWhere + vbCr
message = message + PizzaSize + " Pizza" + vbCr
message = message + PizzaCrust + vbCr
message = message + "Cheese only" + vbCr

For I = 0 To 5
If chkTop(I).Value = vbChecked Then message = message + chkTop(I).Caption + vbCr
Next I
message = message + "TOTAL AMOUNT= " + Text1.Text + vbCr
MsgBox message, vbOKOnly, "YOUR PIZZA"

End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Command3_Click()
Text1.Text = Val(Label1) + Val(Label4) + Val(Label5) + Val(Label6) + Val(Label7) + Val(Label8) + Val(Label9)

End Sub

Private Sub Form_Load()
PizzaSize = "Small"
PizzaCrust = "Thin Crust"
PizzaWhere = "Eat In"
End Sub

Private Sub optCrust_Click(Index As Integer)
PizzaCrust = optCrust(Index).Caption
If optCrust(0).Value = True Then
Label10.Caption = 50
End If
If optCrust(1).Value = True Then
Label10.Caption = 60
End If

End Sub

Private Sub optSize_Click(Index As Integer)
PizzaSize = optSize(Index).Caption

If optSize(0).Value = True Then
Label1.Caption = 30
End If
If optSize(1).Value = True Then
Label1.Caption = 40
End If
If optSize(2).Value = True Then
Label1.Caption = 50
End If
End Sub

Private Sub optWhere_Click(Index As Integer)
PizzaWhere = optWhere(Index).Caption
End Sub

