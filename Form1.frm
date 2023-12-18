VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   15690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   14400
      MultiLine       =   -1  'True
      TabIndex        =   92
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   15720
      TabIndex        =   90
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   1455
      Left            =   10440
      MultiLine       =   -1  'True
      TabIndex        =   89
      Top             =   3000
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   13320
      MultiLine       =   -1  'True
      TabIndex        =   88
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Puzzle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12480
      TabIndex        =   87
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Solve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10440
      TabIndex        =   1
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Fitness Tracker:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   91
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   80
      Left            =   9120
      TabIndex        =   86
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   79
      Left            =   8040
      TabIndex        =   85
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   78
      Left            =   6960
      TabIndex        =   84
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   77
      Left            =   5640
      TabIndex        =   83
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   76
      Left            =   4560
      TabIndex        =   82
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   75
      Left            =   3480
      TabIndex        =   81
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   74
      Left            =   2160
      TabIndex        =   80
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   73
      Left            =   1080
      TabIndex        =   79
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   72
      Left            =   0
      TabIndex        =   78
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   71
      Left            =   9120
      TabIndex        =   77
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   70
      Left            =   8040
      TabIndex        =   76
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   69
      Left            =   6960
      TabIndex        =   75
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   68
      Left            =   5640
      TabIndex        =   74
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   67
      Left            =   4560
      TabIndex        =   73
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   66
      Left            =   3480
      TabIndex        =   72
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   65
      Left            =   2160
      TabIndex        =   71
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   64
      Left            =   1080
      TabIndex        =   70
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   63
      Left            =   0
      TabIndex        =   69
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   62
      Left            =   9120
      TabIndex        =   68
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   61
      Left            =   8040
      TabIndex        =   67
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   60
      Left            =   6960
      TabIndex        =   66
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   59
      Left            =   5640
      TabIndex        =   65
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   58
      Left            =   4560
      TabIndex        =   64
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   57
      Left            =   3480
      TabIndex        =   63
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   56
      Left            =   2160
      TabIndex        =   62
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   55
      Left            =   1080
      TabIndex        =   61
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   54
      Left            =   0
      TabIndex        =   60
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   53
      Left            =   9120
      TabIndex        =   59
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   52
      Left            =   8040
      TabIndex        =   58
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   51
      Left            =   6960
      TabIndex        =   57
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   50
      Left            =   5640
      TabIndex        =   56
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   49
      Left            =   4560
      TabIndex        =   55
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   48
      Left            =   3480
      TabIndex        =   54
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   47
      Left            =   2160
      TabIndex        =   53
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   46
      Left            =   1080
      TabIndex        =   52
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   45
      Left            =   0
      TabIndex        =   51
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   44
      Left            =   9120
      TabIndex        =   50
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   43
      Left            =   8040
      TabIndex        =   49
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   42
      Left            =   6960
      TabIndex        =   48
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   41
      Left            =   5640
      TabIndex        =   47
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   40
      Left            =   4560
      TabIndex        =   46
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   39
      Left            =   3480
      TabIndex        =   45
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   38
      Left            =   2160
      TabIndex        =   44
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   37
      Left            =   1080
      TabIndex        =   43
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   36
      Left            =   0
      TabIndex        =   42
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   35
      Left            =   9120
      TabIndex        =   41
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   34
      Left            =   8040
      TabIndex        =   40
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   33
      Left            =   6960
      TabIndex        =   39
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   32
      Left            =   5640
      TabIndex        =   38
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   31
      Left            =   4560
      TabIndex        =   37
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   30
      Left            =   3480
      TabIndex        =   36
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   29
      Left            =   2160
      TabIndex        =   35
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   28
      Left            =   1080
      TabIndex        =   34
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   27
      Left            =   0
      TabIndex        =   33
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   26
      Left            =   9120
      TabIndex        =   32
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   25
      Left            =   8040
      TabIndex        =   31
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   24
      Left            =   6960
      TabIndex        =   30
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   23
      Left            =   5640
      TabIndex        =   29
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   22
      Left            =   4560
      TabIndex        =   28
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   21
      Left            =   3480
      TabIndex        =   27
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   20
      Left            =   2160
      TabIndex        =   26
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   19
      Left            =   1080
      TabIndex        =   25
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   18
      Left            =   0
      TabIndex        =   24
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   17
      Left            =   9120
      TabIndex        =   23
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   16
      Left            =   8040
      TabIndex        =   22
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   15
      Left            =   6960
      TabIndex        =   21
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   14
      Left            =   5640
      TabIndex        =   20
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   13
      Left            =   4560
      TabIndex        =   19
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   12
      Left            =   3480
      TabIndex        =   18
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   2160
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   1080
      TabIndex        =   16
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   0
      TabIndex        =   15
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   9120
      TabIndex        =   14
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   7
      Left            =   8040
      TabIndex        =   13
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   6960
      TabIndex        =   12
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   5640
      TabIndex        =   11
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   4560
      TabIndex        =   10
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   3480
      TabIndex        =   9
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   2160
      TabIndex        =   8
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Result:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   6
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Final Arrangement:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   5
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Genetic Algorithm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   10440
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Fitness Score: -1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Generation: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* Robert Powell
'* CSCI 4307
'* Ken Nguyen, PhD.
'* 12/01/2022
'*
'* This program is an implementations of a genetic algorithm
'* to solve a Sudoku puzzle.
'*

Public userRows, userColumns, evaluationTimeAtStart, evaluationTimeAtEnd, stuck, maxArraySize



'Start Solve

Private Sub Command1_Click()
geneticAlgorithm
End Sub
'Start Solve end

Sub geneticAlgorithm()
    
    'Determine Number Of Solutions To Generated
    solutionsToGenerate = 80 - determineNumberOfFixedSpaces()
    
    'disable start button
    Command1.Enabled = False
    
    'Set populaiton size.
    populationSize = 5
    
    'Clear the previous population members if there are any.
    populationMembers$ = ""
    
    'Set generation limit.
    generationLimit = 500
    
    'Clear previously displayed information
    Label3.Caption = "Fitness Score: -1"
    Label2.Caption = "Generation: 0"
    Label5.Caption = "Final Arrangement: 0 0 0 0 0 0 0 0"
    Label6.Caption = "Result:"
    
    'Begin the creation of the initial population.
    populationMembers$ = "*"
    For populationCounter = 1 To populationSize
        
        newMember$ = ""
        For generatedSolutionsCounter = 0 To solutionsToGenerate
            Randomize
            randomNumber = generateRandomNumber(1, 9)
            newMember$ = newMember$ + Str$(randomNumber)
         Next generatedSolutionsCounter
        
        newMember$ = LTrim(RTrim(newMember$))
        If InStr(1, populationMembers$, newMember$) = 0 Then
            populationMembers$ = populationMembers$ + "|" + newMember$
            populationTracker = populationTracker + 1
        End If
        
        'If populationCounter <> populationSize Then populationMembers$ = populationMembers$ + "|"
        
    Next populationCounter
    
    populationMembers$ = Replace(populationMembers$, "*|", "")
    populationMembers$ = Replace(populationMembers$, "| ", "|")
    
    'Check the fitness of the initial population members.
    generationCounter = 0
    fitnessResults$ = findMostFit(populationMembers$)

    'Mix population members until a solution is found or the generation limit is reached.
    Do While InStr(1, fitnessResults$, "/0:") = "0" And generationCounter <> generationLimit
        mixResults$ = mix(fitnessResults$)
        populationMembers$ = mixResults$
        fitnessResults$ = findMostFit(populationMembers$)
        generationCounter = generationCounter + 1
        Label2.Caption = "Generation: " + Str(generationCounter)
    Loop
    
    'Display the best solution found the board
    fitnessResultsComponents = Split(fitnessResults$, "|")
    finalArrangementAndFitnessScore = Split(fitnessResultsComponents(0), ":")
    
    Label3.Caption = "Target Value: " + Replace(finalArrangementAndFitnessScore(0), "/", "")
    Label2.Caption = "Generation: " + Str(generationCounter)
    finalOutputString$ = findMostFit(finalArrangementAndFitnessScore(1), True)
    generateFullSolution
    
    If Replace(finalArrangementAndFitnessScore(0), "/", "") = "0" Then
        Label6.Caption = "Result: Success!"
    Else
        Label6.Caption = "Result: Failure..."
    End If
    
    're-enable start button
    Command1.Enabled = True
    Command1.Caption = "Try Again?"

End Sub
'Determine the number of clue spaces present
Function determineNumberOfFixedSpaces()
    foundFixedSpaces = 0
    For spaceCounter = 0 To 80
       If Label1(spaceCounter).Caption <> "" Then foundFixedSpaces = foundFixedSpaces + 1
    Next spaceCounter
    determineNumberOfFixedSpaces = foundFixedSpaces
End Function


'Mix the most fit population member with the rest of the population
Function mix(populationToMix As String)
    
    mostFitInfoComponents = Split(populationToMix, ":")
    
    populationToMixComponents = Split(mostFitInfoComponents(1), "|")

Text3.Text = Str(UBound(populationToMixComponents)) + " " + Text3.Text

    For mostFitCounter = 0 To UBound(populationToMixComponents)
    
        'Mixing process begins here. Two offspring are produced with each pairing.
        For mixingCounter = mostFitCounter To UBound(populationToMixComponents)
            
            DoEvents
            
            If populationToMixComponents(mostFitCounter) <> populationToMixComponents(mixingCounter) Then
                maxSolutionLength = Len(populationToMixComponents(mostFitCounter))
                crossPosition = generateRandomNumber(1, maxSolutionLength - 1)
                firstPartOfMostFit$ = Mid(populationToMixComponents(mostFitCounter), 1, crossPosition)
                secondPartOfMostFit$ = Mid(populationToMixComponents(mostFitCounter), crossPosition + 1, maxSolutionLength - crossPosition)
            
                firstPartMemberToMix$ = Mid(populationToMixComponents(mixingCounter), 1, crossPosition)
                secondPartMemberToMix$ = Mid(populationToMixComponents(mixingCounter), crossPosition + 1, maxSolutionLength - crossPosition)
            
                firstOffspring$ = firstPartMemberToMix$ + secondPartOfMostFit$
                secondOffspring$ = firstPartOfMostFit$ + secondPartMemberToMix$
            
                '15% chance to mutate the offspring
                If generateRandomNumber(1, 100) <= 15 Then
                    For mutationCounter = 1 To 1
                        firstOffspring$ = mutate(firstOffspring$)
                        secondOffspring$ = mutate(secondOffspring$)
                    Next mutationCounter
                End If
            
                'Add newly produced offspring to the new population
                If InStr(1, newPopulation$, firstOffspring$) = 0 Then
                    newPopulation$ = newPopulation$ + firstOffspring$ + "|"
                End If
                
                If InStr(1, newPopulation$, secondOffspring$) = 0 Then
                    newPopulation$ = newPopulation$ + secondOffspring$ + "|"
                End If
                
            End If
        Next mixingCounter
    
    Next mostFitCounter
    
    mix = Mid(newPopulation$, 1, Len(newPopulation$) - 1)
    
End Function

'Randomly modify 1 position within the offspring's arrangment
Function mutate(offSpring As String)
    LengthOfOffspring = Len(offSpring)
    trueLengthOfOffspring = Len(Replace(offSpring, " ", ""))
    mutationPosition = generateRandomNumber(1, Int(trueLengthOfOffspring))
    If mutationPosition = 1 Then
        mutate = LTrim(RTrim(Str(generateRandomNumber(1, 9)))) + Mid(offSpring$, 2, LengthOfOffspring - 1)
    ElseIf mutationPosition = trueLengthOfOffspring Then
        mutate = Mid(offSpring, 1, LengthOfOffspring - 1) + LTrim(RTrim(Str(generateRandomNumber(1, 9))))
    Else
        mutate = Mid(offSpring$, 1, 2 * mutationPosition - 2) + LTrim(Str(generateRandomNumber(1, 9))) + " " + LTrim(RTrim(Mid(offSpring, 2 * mutationPosition, LengthOfOffspring - (2 * mutationPosition - 1))))
    End If
End Function

'Random number generator
Function generateRandomNumber(lowerBound As Integer, upperBound As Integer)
Randomize
generateRandomNumber = Int((upperBound - lowerBound + 1) * Rnd + lowerBound)
End Function

'clear the board
Sub clearBoard()
    For evaluationClearCounter = 0 To Label1.UBound
        If Label1(evaluationClearCounter).Tag <> "1" Then
            Label1(evaluationClearCounter).Caption = ""
            Label1(evaluationClearCounter).BackColor = &H8000000F
        Else:
            Label1(evaluationClearCounter).BackColor = vbBlue
        End If
    Next evaluationClearCounter
End Sub

'check columns for violations
Function checkColumns(callMarkSquares, boardValuesToCheck)
    For columnIndex = 0 To 8
        DoEvents
        numbersToCheckFor$ = "123456789"
        For columnEvaluationCounter = columnIndex To columnIndex + 72 Step 9
            currentValueInColumn$ = boardValuesToCheck(columnEvaluationCounter)
            If InStr(1, numbersToCheckFor$, currentValueInColumn$) <> 0 Then
                numbersToCheckFor$ = Replace(numbersToCheckFor$, currentValueInColumn$, "")
            Else
                If callMarkSquares = True Then markSquares (columnEvaluationCounter)
                If numbersToCheckFor$ <> "" Then fitnessDemerits = fitnessDemerits + 1
            End If
        Next columnEvaluationCounter
    checkColumns = fitnessDemerits
    Next columnIndex
End Function

'check rows for violations
Function checkRows(callMarkSquares, boardValuesToCheck)
    For rowIndex = 0 To 72 Step 9
        DoEvents
        numbersToCheckFor$ = "123456789"
        For rowEvaluationCounter = rowIndex To rowIndex + 8
            currentValueInColumn$ = boardValuesToCheck(rowEvaluationCounter)
            If InStr(1, numbersToCheckFor$, currentValueInColumn$) <> 0 Then
                numbersToCheckFor$ = Replace(numbersToCheckFor$, currentValueInColumn$, "")
            Else
                If callMarkSquares = True Then markSquares (rowEvaluationCounter)
                If numbersToCheckFor$ <> "" Then fitnessDemerits = fitnessDemerits + 1
            End If
        Next rowEvaluationCounter
    checkRows = fitnessDemerits
    Next rowIndex
End Function

'check subgrids for violations
Function checkSubGrids(callMarkSquares, boardValuesToCheck)
    For rowIndex = 0 To 72 Step 27
        DoEvents
        For gridCounter = rowIndex To rowIndex + 8 Step 3
            numbersToCheckFor$ = "123456789"
            For columnCounter = gridCounter To gridCounter + 2
                For columnIndex = columnCounter To columnCounter + 18 Step 9
                    currentValueInColumn$ = boardValuesToCheck(columnIndex)
                    If InStr(1, numbersToCheckFor$, currentValueInColumn$) <> 0 Then
                        numbersToCheckFor$ = Replace(numbersToCheckFor$, currentValueInColumn$, "")
                    Else
                        If callMarkSquares = True Then markSquares (columnIndex)
                        If numbersToCheckFor$ <> "" Then fitnessDemerits = fitnessDemerits + 1
                    End If
                Next columnIndex
            Next columnCounter
        Next gridCounter
    Next rowIndex
    checkSubGrids = fitnessDemerits

End Function

'analyze board
Function compAnaylsis(markViolations, boardValues)
    compAnaylsis = checkSubGrids(markViolations, boardValues) + checkRows(markViolations, boardValues) + checkColumns(markViolations, boardValues)
End Function

'Button Loads soduku puzzle
Private Sub Command2_Click()
    LoadPuzzle ("700000000901003065020158400000009008302046000090325000030000010048200970070064030")
End Sub

'Loads soduku puzzle
Sub LoadPuzzle(puzzle As String)

    For puzzleCounter = 1 To 81
        If Mid$(puzzle, puzzleCounter, 1) <> "0" Then
            Label1(puzzleCounter - 1).BackColor = vbBlue
            Label1(puzzleCounter - 1).ForeColor = vbYellow
            Label1(puzzleCounter - 1).Caption = Mid(puzzle, puzzleCounter, 1)
            Label1(puzzleCounter - 1).Tag = "1"
        End If
    Next puzzleCounter

End Sub

Private Sub Form_Load()

   'Board Specifications
    userColumns = 9
    userRows = 9
    
End Sub

'Analyzes population for the most fit arrangement
Function findMostFit(populationRepresentation, Optional displayOveride As Boolean)
    
    Dim board(80) As String
    
    'Set population array sorting variables
    Dim allPopulationMembers(200) As String
    
    If stuck < 10 Then
        maxArraySize = 25
    ElseIf maxArraySize <> 50 Then
        maxArraySize = maxArraySize + 5
    End If
    
    numberOfPopulationMembersRanked = 0
    
    'Set standards for fitness and culling
    mostFitValue = 1000
    If evaluationTimeAtStart = 0 Then evaluationTimeAtStart = -1
    Text1.Text = ""
    
    'Search for most fit arrangement
    populationComponents = Split(populationRepresentation, "|")
    For populationComponentsIndex = 0 To UBound(populationComponents)
        
        DoEvents
        
        individualMember = Split(populationComponents(populationComponentsIndex), " ")
        
        If evaluationTimeAtStart <> -1 Then evaluationTimeAtEnd = Timer

        If Abs(evaluationTimeAtStart - evaluationTimeAtEnd) >= 0.5 Or displayOveride = True Then
            evaluationTimeAtStart = -1
            displayArrangement = True
            Call clearBoard
        Else
            displayArrangement = False
        End If
        
        'Place potential solution in array and evaluate its fitness
        individualMemberCounter = 0
        For evaluationCounter = 0 To 80
            
            If Label1(evaluationCounter).Tag <> "1" And displayArrangement = True Then
                Label1(evaluationCounter).Caption = individualMember(individualMemberCounter)
            End If
                
            If Label1(evaluationCounter).Tag <> "1" Then
                board(evaluationCounter) = individualMember(individualMemberCounter)
                individualMemberCounter = individualMemberCounter + 1
            ElseIf Label1(evaluationCounter).Tag = "1" Then
                board(evaluationCounter) = Label1(evaluationCounter).Caption
            End If
            
        Next evaluationCounter
        fitnessValue = compAnaylsis(displayArrangement, board)
        
        Text1.Text = Str(fitnessValue) + " " + Text1.Text
        
        'Sorts Population members in array from lowest fitness demerits to highest.
        If numberOfPopulationMembersRanked = 0 Then
            allPopulationMembers(0) = populationComponents(populationComponentsIndex) + ":" + LTrim(RTrim(Str(fitnessValue)))
            numberOfPopulationMembersRanked = numberOfPopulationMembersRanked + 1
        Else
            rankedPopulationMembersCounter = 0
            greaterNumberFound = 0
            Do While rankedPopulationMembersCounter <> numberOfPopulationMembersRanked And greaterNumberFound = 0
                allPopulationMembersComponents = Split(allPopulationMembers(rankedPopulationMembersCounter), ":")
                If fitnessValue <= Val(allPopulationMembersComponents(1)) Then
                    greaterNumberFound = 1
                Else
                    rankedPopulationMembersCounter = rankedPopulationMembersCounter + 1
                End If
            Loop
            
            populationMemberToInsert$ = populationComponents(populationComponentsIndex) + ":" + LTrim(RTrim(Str(fitnessValue)))
            
            If greaterNumberFound = 0 And maxArraySize + 1 <> numberOfPopulationMembersRanked Then
                allPopulationMembers(numberOfPopulationMembersRanked) = populationMemberToInsert$
                numberOfPopulationMembersRanked = numberOfPopulationMembersRanked + 1
                greaterNumberFound = 0
            ElseIf greaterNumberFound = 1 And maxArraySize + 1 <> numberOfPopulationMembersRanked Then
                sortStartPoint = numberOfPopulationMembersRanked
                numberOfPopulationMembersRanked = numberOfPopulationMembersRanked + 1
            ElseIf greaterNumberFound = 1 And maxArraySize + 1 = numberOfPopulationMembersRanked Then
                sortStartPoint = numberOfPopulationMembersRanked - 1
            End If
            
            sortEndIndex = rankedPopulationMembersCounter + 1
            If greaterNumberFound = 1 Then
                For sortCounter = sortStartPoint To sortEndIndex Step -1
                    allPopulationMembers(sortCounter) = allPopulationMembers(sortCounter - 1)
                Next sortCounter
                allPopulationMembers(rankedPopulationMembersCounter) = populationMemberToInsert$
                greaterNumberFound = 0
            End If
            
        End If
        
        
        If evaluationTimeAtStart = -1 Then evaluationTimeAtStart = Timer
        
        'Compare fitness against other arrangements
        If fitnessValue < mostFitValue Then
            mostFitValue = fitnessValue
        End If
    
    Next populationComponentsIndex
    
    For stringBuilderCounter = 0 To numberOfPopulationMembersRanked - 1
        allPopulationMembersComponents = Split(allPopulationMembers(stringBuilderCounter), ":")
        If stringBuilderCounter = 0 Then
            outputString$ = "/" + allPopulationMembersComponents(1) + ":" + allPopulationMembersComponents(0)
        ElseIf stringBuilderCounter > 0 Then
            outputString$ = outputString$ + "|" + allPopulationMembersComponents(0)
        End If
    Next stringBuilderCounter

    findMostFit = outputString$
    Label3.Caption = "Fitness Score: " + Str(mostFitValue)
    
    If Label3.Tag = Str(mostFitValue) Then
        stuck = stuck + 1
    Else:
        Label3.Tag = Str(mostFitValue)
        stuck = 0
        maxArraySize = 25
    End If
    
End Function

'Mark rule violating squares
Sub markSquares(boardPosition)
    Label1(boardPosition).BackColor = vbRed
End Sub

'Debugging Code
Private Sub Label1_DblClick(index As Integer)

    If Val(Label1(index).Caption) < 9 Then
        Label1(index).Caption = Str$(Val(Val(Label1(index).Caption)) + 1)
        Label1(index).BackColor = vbBlue
        Label1(index).ForeColor = vbYellow
        Label1(index).Tag = "1"
    Else
        Label1(index).Caption = ""
        Label1(index).BackColor = &H8000000F
        Label1(index).ForeColor = vbBlack
        Label1(index).DataMember = 0
    End If

End Sub

'Generate Solution Located After Quiting Search
Sub generateFullSolution()
    Text2.Text = ""
    For generateSolutionCounter = 0 To 80
        Text2.Text = Text2.Text + Label1(generateSolutionCounter).Caption
    Next
End Sub
