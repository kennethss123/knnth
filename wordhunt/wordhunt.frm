VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0080FFFF&
   Caption         =   "MISTAKES"
   ClientHeight    =   6300
   ClientLeft      =   6480
   ClientTop       =   4215
   ClientWidth     =   15885
   LinkTopic       =   "Form3"
   ScaleHeight     =   6300
   ScaleWidth      =   15885
   Begin VB.TextBox Text12 
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
      Left            =   14160
      TabIndex        =   124
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text11 
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
      Left            =   11640
      TabIndex        =   122
      Text            =   "10"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Text10 
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
      Left            =   13800
      TabIndex        =   115
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text9 
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
      Left            =   13800
      TabIndex        =   114
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text8 
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
      Left            =   13800
      TabIndex        =   113
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text7 
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
      Left            =   13800
      TabIndex        =   112
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text6 
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
      Left            =   13800
      TabIndex        =   111
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text5 
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
      Left            =   11040
      TabIndex        =   110
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text4 
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
      Left            =   11040
      TabIndex        =   109
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
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
      Left            =   11040
      TabIndex        =   108
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Left            =   11040
      TabIndex        =   107
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   104
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
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
      Left            =   11040
      TabIndex        =   101
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command100 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   98
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command99 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   97
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command98 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   96
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command97 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   95
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command96 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   94
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command95 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   93
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command94 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   92
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command93 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   91
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command92 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   90
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command91 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   89
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command90 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   88
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command89 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   87
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command88 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   86
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command87 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   85
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command86 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   84
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command85 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   83
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command84 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   82
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command83 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   81
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command82 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   80
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command81 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   79
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command80 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   78
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command79 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   77
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command78 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   76
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command77 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   75
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command76 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   74
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command75 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   73
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command74 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   72
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command73 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   71
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command72 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   70
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command71 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   69
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command70 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   68
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command69 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   67
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command68 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   66
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command67 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   65
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command66 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   64
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command65 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   63
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command64 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   62
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command63 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   61
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command62 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   60
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command61 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   59
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command60 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   58
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command59 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   57
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command58 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   56
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command57 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   55
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command56 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   54
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command55 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   53
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command54 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   52
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command53 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   51
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command52 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   50
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command51 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   49
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command50 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   48
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command49 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   47
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command48 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   46
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command47 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   45
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command46 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   44
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command45 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   43
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command44 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   42
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command43 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   41
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command42 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   40
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command41 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   39
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command40 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   38
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command39 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   37
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command38 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   36
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command37 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   35
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command36 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   34
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command35 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   33
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command34 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   32
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command33 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   31
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command32 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   30
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command31 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   29
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command30 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   28
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command29 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   27
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command28 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   26
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command27 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   25
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command26 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   24
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command25 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   23
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command24 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   22
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command23 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   21
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command22 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   20
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command21 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   19
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command20 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   18
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command19 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command18 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   16
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   15
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "DEVELOPED BY: TEAM RED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11520
      TabIndex        =   125
      Top             =   5880
      Width           =   5655
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "MISTAKES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   123
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "REMAINING WORDS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9720
      TabIndex        =   121
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "MOUSE"
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
      Left            =   12480
      TabIndex        =   120
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "POWER"
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
      Left            =   12480
      TabIndex        =   119
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "FAN"
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
      Left            =   12480
      TabIndex        =   118
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "CABLE"
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
      Left            =   12480
      TabIndex        =   117
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "MEMES"
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
      Left            =   12480
      TabIndex        =   116
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "MONITOR"
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
      Left            =   9720
      TabIndex        =   106
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "TELEPHONE"
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
      Left            =   9720
      TabIndex        =   105
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "INTERNET"
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
      Left            =   9720
      TabIndex        =   103
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "WIRE"
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
      Left            =   9720
      TabIndex        =   102
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "CELLPHONE"
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
      Left            =   9720
      TabIndex        =   100
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WORDS TO FIND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9720
      TabIndex        =   99
      Top             =   360
      Width           =   6615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.BackColor = vbRed
Command1.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Text1.Text = 9 Then
Text11.Text = Val(Text11.Text) - 1
Label2.Visible = False
Text1.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If



End Sub

Private Sub Command10_Click()
Command10.BackColor = vbRed
Command10.Enabled = False
Text7.Text = Val(Text7.Text) + 1
If Text7.Text = 5 Then
Text11.Text = Val(Text11.Text) - 1
Label8.Visible = False
Text7.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command100_Click()
Command100.BackColor = vbRed
Command100.Enabled = False
Text4.Text = Val(Text4.Text) + 1
If Text4.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label5.Visible = False
Text4.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command11_Click()
MsgBox "Oops Wrong"
Command11.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command12_Click()
MsgBox "Oops Wrong"
Command12.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command13_Click()
MsgBox "Oops Wrong"
Command13.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command14_Click()
Command14.BackColor = vbRed
Command14.Enabled = False
Text8.Text = Val(Text8.Text) + 1
If Text8.Text = 3 Then
Text11.Text = Val(Text11.Text) - 1
Label9.Visible = False
Text8.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command15_Click()
MsgBox "Oops Wrong"
Command15.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command16_Click()
MsgBox "Oops Wrong"
Command16.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command17_Click()
MsgBox "Oops Wrong"
Command17.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command18_Click()
Command18.BackColor = vbRed
Command18.Enabled = False
Text4.Text = Val(Text4.Text) + 1
If Text4.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label5.Visible = False
Text4.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command19_Click()
Command19.BackColor = vbRed
Command19.Enabled = False
Text7.Text = Val(Text7.Text) + 1
If Text7.Text = 5 Then
Text11.Text = Val(Text11.Text) - 1
Label8.Visible = False
Text7.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command2_Click()
Command2.BackColor = vbRed
Command2.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Text1.Text = 9 Then
Text11.Text = Val(Text11.Text) - 1
Label2.Visible = False
Text1.Visible = False

End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If

End Sub

Private Sub Command20_Click()
MsgBox "Oops Wrong"
Command20.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command21_Click()
Command21.BackColor = vbRed
Command21.Enabled = False
Text10.Text = Val(Text10.Text) + 1
If Text10.Text = 5 Then
Text11.Text = Val(Text11.Text) - 1
Label11.Visible = False
Text10.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command22_Click()
MsgBox "Oops Wrong"
Command22.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command23_Click()
Command23.BackColor = vbRed
Command23.Enabled = False
Text4.Text = Val(Text4.Text) + 1
If Text4.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label5.Visible = False
Text4.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command24_Click()
MsgBox "Oops Wrong"
Command24.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command25_Click()
Command25.BackColor = vbRed
Command25.Enabled = False
Text8.Text = Val(Text8.Text) + 1
If Text8.Text = 3 Then
Text1.Text = Val(Text11.Text) - 1
Label9.Visible = False
Text8.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command26_Click()
Command26.BackColor = vbRed
Command26.Enabled = False
Text3.Text = Val(Text3.Text) + 1
If Text3.Text = 7 Then
Text11.Text = Val(Text11.Text) - 1
Label6.Visible = False
Text3.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command27_Click()
MsgBox "Oops Wrong"
Command27.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command28_Click()
MsgBox "Oops Wrong"
Command28.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command29_Click()
Command29.BackColor = vbRed
Command29.Enabled = False
Text7.Text = Val(Text7.Text) + 1
If Text7.Text = 5 Then
Text11.Text = Val(Text11.Text) - 1
Label8.Visible = False
Text7.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command3_Click()
Command3.BackColor = vbRed
Command3.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Text1.Text = 9 Then
Text11.Text = Val(Text11.Text) - 1
Label2.Visible = False
Text1.Visible = False

End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command30_Click()
Command30.BackColor = vbRed
Command30.Enabled = False
Text10.Text = Val(Text10.Text) + 1
If Text10.Text = 5 Then
Text11.Text = Val(Text11.Text) - 1
Label11.Visible = False
Text10.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command31_Click()
MsgBox "Oops Wrong"
Command31.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command32_Click()
Command32.BackColor = vbRed
Command32.Enabled = False
Text4.Text = Val(Text4.Text) + 1
If Text4.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label5.Visible = False
Text4.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command33_Click()
MsgBox "Oops Wrong"
Command33.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command34_Click()
MsgBox "Oops Wrong"
Command34.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command35_Click()
Command35.BackColor = vbRed
Command35.Enabled = False
Text8.Text = Val(Text8.Text) + 1
If Text8.Text = 3 Then
Text1.Text = Val(Text11.Text) - 1
Label9.Visible = False
Text8.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command36_Click()
MsgBox "Oops Wrong"
Command36.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command37_Click()
MsgBox "Oops Wrong"
Command37.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command38_Click()
MsgBox "Oops Wrong"
Command38.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command39_Click()
Command39.BackColor = vbRed
Command39.Enabled = False
Text7.Text = Val(Text7.Text) + 1
If Text7.Text = 5 Then
Text11.Text = Val(Text11.Text) - 1
Label8.Visible = False
Text7.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command4_Click()
Command4BackColor = vbRed
Command4.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Text1.Text = 9 Then
Text11.Text = Val(Text11.Text) - 1
Label2.Visible = False
Text1.Visible = False

End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If

End Sub

Private Sub Command40_Click()
MsgBox "Oops Wrong"
Command40.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command41_Click()
Command41.BackColor = vbRed
Command41.Enabled = False
Text7.Text = Val(Text7.Text) + 1
If Text7.Text = 5 Then
Text11.Text = Val(Text11.Text) - 1
Label8.Visible = False
Text7.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command42_Click()
MsgBox "Oops Wrong"
Command42.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If


End Sub

Private Sub Command43_Click()
MsgBox "Oops Wrong"
Command43.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command44_Click()
Command44.BackColor = vbRed
Command44.Enabled = False
Text10.Text = Val(Text10.Text) + 1
If Text10.Text = 5 Then
Text11.Text = Val(Text11.Text) - 1
Label11.Visible = False
Text10.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command45_Click()
MsgBox "Oops Wrong"
Command45.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command46_Click()
Command46.BackColor = vbRed
Command46.Enabled = False
Text4.Text = Val(Text4.Text) + 1
If Text4.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label5.Visible = False
Text4.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command47_Click()
Command47.BackColor = vbRed
Command47.Enabled = False
Text3.Text = Val(Text3.Text) + 1
If Text3.Text = 7 Then
Text11.Text = Val(Text11.Text) - 1
Label6.Visible = False
Text3.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command48_Click()
Command48.BackColor = vbRed
Command48.Enabled = False
Text9.Text = Val(Text9.Text) + 1
If Text9.Text = 3 Then
Text1.Text = Val(Text11.Text) - 1
Label10.Visible = False
Text9.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command49_Click()
Command49.BackColor = vbRed
Command49.Enabled = False
Text9.Text = Val(Text9.Text) + 1
If Text9.Text = 3 Then
Text1.Text = Val(Text11.Text) - 1
Label10.Visible = False
Text9.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command5_Click()
Command5.BackColor = vbRed
Command5.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Text1.Text = 9 Then
Text11.Text = Val(Text11.Text) - 1
Label2.Visible = False
Text1.Visible = False

End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command50_Click()
Command50.BackColor = vbRed
Command50.Enabled = False
Text9.Text = Val(Text9.Text) + 1
If Text9.Text = 3 Then
Text11.Text = Val(Text11.Text) - 1
Label10.Visible = False
Text9.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command51_Click()
MsgBox "Oops Wrong"
Command51.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command52_Click()
Command52.BackColor = vbRed
Command52.Enabled = False
Text6.Text = Val(Text6.Text) + 1
If Text6.Text = 4 Then
Text11.Text = Val(Text11.Text) - 1
Label7.Visible = False
Text6.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command53_Click()
MsgBox "Oops Wrong"
Command53.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command54_Click()
MsgBox "Oops Wrong"
Command54.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command55_Click()
Command55.BackColor = vbRed
Command55.Enabled = False
Text10.Text = Val(Text10.Text) + 1
If Text10.Text = 5 Then
Text11.Text = Val(Text11.Text) - 1
Label11.Visible = False
Text10.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command56_Click()
Command56.BackColor = vbRed
Command56.Enabled = False
Text3.Text = Val(Text3.Text) + 1
If Text3.Text = 7 Then
Text11.Text = Val(Text11.Text) - 1
Label6.Visible = False
Text3.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command57_Click()
Command57.BackColor = vbRed
Command57.Enabled = False
Text4.Text = Val(Text4.Text) + 1
If Text4.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label5.Visible = False
Text4.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command58_Click()
MsgBox "Oops Wrong"
Command58.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command59_Click()
MsgBox "Oops Wrong"
Command59.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command6_Click()
Command6.BackColor = vbRed
Command6.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Text1.Text = 9 Then
Text11.Text = Val(Text11.Text) - 1
Label2.Visible = False
Text1.Visible = False

End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command60_Click()
MsgBox "Oops Wrong"
Command60.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command61_Click()
Command61.BackColor = vbRed
Command61.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Text2.Text = 4 Then
Text11.Text = Val(Text11.Text) - 1
Label3.Visible = False
Text2.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If

End Sub

Private Sub Command62_Click()
Command62.BackColor = vbRed
Command62.Enabled = False
Text6.Text = Val(Text6.Text) + 1
If Text6.Text = 4 Then
Text11.Text = Val(Text11.Text) - 1
Label7.Visible = False
Text6.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command63_Click()
MsgBox "Oops Wrong"
Command63.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command64_Click()
MsgBox "Oops Wrong"
Command64.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command65_Click()
Command65.BackColor = vbRed
Command65.Enabled = False
Text3.Text = Val(Text3.Text) + 1
If Text3.Text = 7 Then
Text11.Text = Val(Text11.Text) - 1
Label6.Visible = False
Text3.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command66_Click()
Command66.BackColor = vbRed
Command66.Enabled = False
Text10.Text = Val(Text10.Text) + 1
If Text10.Text = 5 Then
Text11.Text = Val(Text11.Text) - 1
Label11.Visible = False
Text10.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command67_Click()
MsgBox "Oops Wrong"
Command67.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command68_Click()
Command68.BackColor = vbRed
Command68.Enabled = False
Text4.Text = Val(Text4.Text) + 1
If Text4.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label5.Visible = False
Text4.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command69_Click()
MsgBox "Oops Wrong"
Command69.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command7_Click()
Command7.BackColor = vbRed
Command7.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Text1.Text = 9 Then
Text11.Text = Val(Text11.Text) - 1
Label2.Visible = False
Text1.Visible = False

End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command70_Click()
MsgBox "Oops Wrong"
Command70.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command71_Click()
Command71.BackColor = vbRed
Command71.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Text2.Text = 4 Then
Text11.Text = Val(Text11.Text) - 1
Label3.Visible = False
Text2.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If

End Sub

Private Sub Command72_Click()
Command72.BackColor = vbRed
Command72.Enabled = False
Text6.Text = Val(Text6.Text) + 1
If Text6.Text = 4 Then
Text11.Text = Val(Text11.Text) - 1
Label7.Visible = False
Text6.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command73_Click()
MsgBox "Oops Wrong"
Command73.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command74_Click()
Command74.BackColor = vbRed
Command74.Enabled = False
Text3.Text = Val(Text3.Text) + 1
If Text3.Text = 7 Then
Text11.Text = Val(Text11.Text) - 1
Label6.Visible = False
Text3.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command75_Click()
MsgBox "Oops Wrong"
Command75.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command76_Click()
MsgBox "Oops Wrong"
Command76.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command77_Click()
MsgBox "Oops Wrong"
Command77.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command78_Click()
MsgBox "Oops Wrong"
Command78.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command79_Click()
Command79.BackColor = vbRed
Command79.Enabled = False
Text4.Text = Val(Text4.Text) + 1
If Text4.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label5.Visible = False
Text4.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command8_Click()
Command8.BackColor = vbRed
Command8.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Text1.Text = 9 Then
Text11.Text = Val(Text11.Text) - 1
Label2.Visible = False
Text1.Visible = False

End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command80_Click()
MsgBox "Oops Wrong"
Command80.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command81_Click()
Command81.BackColor = vbRed
Command81.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Text2.Text = 4 Then
Text11.Text = Val(Text11.Text) - 1
Label3.Visible = False
Text2.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If

End Sub

Private Sub Command82_Click()
Command82.BackColor = vbRed
Command82.Enabled = False
Text3.Text = Val(Text3.Text) + 1
If Text3.Text = 7 Then
Text11.Text = Val(Text11.Text) - 1
Label6.Visible = False
Text3.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command83_Click()
Command83.BackColor = vbRed
Command83.Enabled = False
Text5.Text = Val(Text5.Text) + 1
If Text5.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label4.Visible = False
Text5.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command84_Click()
Command84.BackColor = vbRed
Command84.Enabled = False
Text5.Text = Val(Text5.Text) + 1
If Text5.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label4.Visible = False
Text5.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command85_Click()
Command85.BackColor = vbRed
Command85.Enabled = False
Text5.Text = Val(Text5.Text) + 1
If Text5.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label4.Visible = False
Text5.Visible = False

End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command86_Click()
Command86.BackColor = vbRed
Command86.Enabled = False
Text5.Text = Val(Text5.Text) + 1
If Text5.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label4.Visible = False
Text5.Visible = False

End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command87_Click()
Command87.BackColor = vbRed
Command87.Enabled = False
Text5.Text = Val(Text5.Text) + 1
If Text5.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label4.Visible = False
Text5.Visible = False

End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command88_Click()
Command88.BackColor = vbRed
Command88.Enabled = False
Text5.Text = Val(Text5.Text) + 1
If Text5.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label4.Visible = False
Text5.Visible = False

End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command89_Click()
Command89.BackColor = vbRed
Command89.Enabled = False
Text5.Text = Val(Text5.Text) + 1
If Text5.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label4.Visible = False
Text5.Visible = False

End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command9_Click()
Command9.BackColor = vbRed
Command9.Enabled = False
Text1.Text = Val(Text1.Text) + 1
If Text1.Text = 9 Then
Text11.Text = Val(Text11.Text) - 1
Label2.Visible = False
Text1.Visible = False



End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If


End Sub

Private Sub Command90_Click()
Command90.BackColor = vbRed
Command90.Enabled = False
Text5.Text = Val(Text5.Text) + 1
If Text5.Text = 8 Then
Text11.Text = Val(Text11.Text) - 1
Label4.Visible = False
Text5.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command91_Click()
Command91.BackColor = vbRed
Command91.Enabled = False
Text2.Text = Val(Text2.Text) + 1
If Text2.Text = 4 Then
Text11.Text = Val(Text11.Text) - 1
Label3.Visible = False
Text2.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If

End Sub

Private Sub Command92_Click()
Command92.BackColor = vbRed
Command92.Enabled = False
Text6.Text = Val(Text6.Text) + 1
If Text6.Text = 4 Then
Text11.Text = Val(Text11.Text) - 1
Label7.Visible = False
Text6.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command93_Click()
Command93.BackColor = vbRed
Command93.Enabled = False
Text3.Text = Val(Text3.Text) + 1
If Text3.Text = 7 Then
Text11.Text = Val(Text11.Text) - 1
Label6.Visible = False
Text3.Visible = False
End If
If Text11.Text = "0" Then
MsgBox "Congrats!"
Form1.Show
Unload Me
End If
End Sub

Private Sub Command94_Click()
MsgBox "Oops Wrong"
Command94.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command95_Click()
MsgBox "Oops Wrong"
Command95.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command96_Click()
MsgBox "Oops Wrong"
Command96.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command97_Click()
MsgBox "Oops Wrong"
Command97.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command98_Click()
MsgBox "Oops Wrong"
Command98.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

Private Sub Command99_Click()
MsgBox "Oops Wrong"
Command99.Enabled = False
Text12.Text = Val(Text12.Text) + 1
If Val(Text12.Text) = 5 Then
Me.Hide
End If
End Sub

