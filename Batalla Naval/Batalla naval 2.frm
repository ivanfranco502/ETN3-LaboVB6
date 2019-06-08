VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnEnd 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5400
      TabIndex        =   102
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton btnGenerar 
      BackColor       =   &H8000000A&
      Caption         =   "Empezar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   5160
      TabIndex        =   101
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   20
      Left            =   120
      TabIndex        =   100
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   6
      Left            =   3000
      TabIndex        =   99
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   7
      Left            =   3480
      TabIndex        =   98
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   8
      Left            =   3960
      TabIndex        =   97
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   9
      Left            =   4440
      TabIndex        =   96
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   95
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   99
      Left            =   4440
      TabIndex        =   93
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   98
      Left            =   3960
      TabIndex        =   92
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   97
      Left            =   3480
      TabIndex        =   91
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   96
      Left            =   3000
      TabIndex        =   90
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   95
      Left            =   2520
      TabIndex        =   89
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   94
      Left            =   2040
      TabIndex        =   88
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   93
      Left            =   1560
      TabIndex        =   87
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   92
      Left            =   1080
      TabIndex        =   86
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   91
      Left            =   600
      TabIndex        =   85
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   90
      Left            =   120
      TabIndex        =   84
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   89
      Left            =   4440
      TabIndex        =   83
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   88
      Left            =   3960
      TabIndex        =   82
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   87
      Left            =   3480
      TabIndex        =   81
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   86
      Left            =   3000
      TabIndex        =   80
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   85
      Left            =   2520
      TabIndex        =   79
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   84
      Left            =   2040
      TabIndex        =   78
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   83
      Left            =   1560
      TabIndex        =   77
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   82
      Left            =   1080
      TabIndex        =   76
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   81
      Left            =   600
      TabIndex        =   75
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   80
      Left            =   120
      TabIndex        =   74
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   79
      Left            =   4440
      TabIndex        =   73
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   78
      Left            =   3960
      TabIndex        =   72
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   77
      Left            =   3480
      TabIndex        =   71
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   76
      Left            =   3000
      TabIndex        =   70
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   75
      Left            =   2520
      TabIndex        =   69
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   74
      Left            =   2040
      TabIndex        =   68
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   73
      Left            =   1560
      TabIndex        =   67
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   72
      Left            =   1080
      TabIndex        =   66
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   71
      Left            =   600
      TabIndex        =   65
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   70
      Left            =   120
      TabIndex        =   64
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   69
      Left            =   4440
      TabIndex        =   63
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   68
      Left            =   3960
      TabIndex        =   62
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   67
      Left            =   3480
      TabIndex        =   61
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   66
      Left            =   3000
      TabIndex        =   60
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   65
      Left            =   2520
      TabIndex        =   59
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   64
      Left            =   2040
      TabIndex        =   58
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   63
      Left            =   1560
      TabIndex        =   57
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   62
      Left            =   1080
      TabIndex        =   56
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   61
      Left            =   600
      TabIndex        =   55
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   60
      Left            =   120
      TabIndex        =   54
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   59
      Left            =   4440
      TabIndex        =   53
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   58
      Left            =   3960
      TabIndex        =   52
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   57
      Left            =   3480
      TabIndex        =   51
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   56
      Left            =   3000
      TabIndex        =   50
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   55
      Left            =   2520
      TabIndex        =   49
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   54
      Left            =   2040
      TabIndex        =   48
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   53
      Left            =   1560
      TabIndex        =   47
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   52
      Left            =   1080
      TabIndex        =   46
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   51
      Left            =   600
      TabIndex        =   45
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   50
      Left            =   120
      TabIndex        =   44
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   49
      Left            =   4440
      TabIndex        =   43
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   48
      Left            =   3960
      TabIndex        =   42
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   47
      Left            =   3480
      TabIndex        =   41
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   46
      Left            =   3000
      TabIndex        =   40
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   45
      Left            =   2520
      TabIndex        =   39
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   44
      Left            =   2040
      TabIndex        =   38
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   43
      Left            =   1560
      TabIndex        =   37
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   42
      Left            =   1080
      TabIndex        =   36
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   41
      Left            =   600
      TabIndex        =   35
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   40
      Left            =   120
      TabIndex        =   34
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   39
      Left            =   4440
      TabIndex        =   33
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   38
      Left            =   3960
      TabIndex        =   32
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   37
      Left            =   3480
      TabIndex        =   31
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   36
      Left            =   3000
      TabIndex        =   30
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   35
      Left            =   2520
      TabIndex        =   29
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   34
      Left            =   2040
      TabIndex        =   28
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   33
      Left            =   1560
      TabIndex        =   27
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   32
      Left            =   1080
      TabIndex        =   26
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   31
      Left            =   600
      TabIndex        =   25
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   30
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   29
      Left            =   4440
      TabIndex        =   23
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   28
      Left            =   3960
      TabIndex        =   22
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   27
      Left            =   3480
      TabIndex        =   21
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   26
      Left            =   3000
      TabIndex        =   20
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   25
      Left            =   2520
      TabIndex        =   19
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   24
      Left            =   2040
      TabIndex        =   18
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   23
      Left            =   1560
      TabIndex        =   17
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   22
      Left            =   1080
      TabIndex        =   16
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   21
      Left            =   600
      TabIndex        =   15
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   19
      Left            =   4440
      TabIndex        =   14
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   18
      Left            =   3960
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   17
      Left            =   3480
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   16
      Left            =   3000
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   15
      Left            =   2520
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   14
      Left            =   2040
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   13
      Left            =   1560
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   12
      Left            =   1080
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   11
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   10
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   5
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   4
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   3
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   2
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Agua 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function SePuede(ByVal fila As Integer, ByVal columna As Integer, matriz() As Byte, ByVal vector As Byte) As Boolean
Dim bytcontador As Byte
Dim bytR As Integer
Dim bytp As Integer

For bytR = fila - 1 To fila + 1
    For bytp = columna - 1 To columna + vector
        If bytR >= 0 And bytR <= 9 Then
            If bytp >= 0 And bytp <= 9 Then
                bytcontador = bytcontador + matriz(bytR, bytp)
            End If
        End If
    Next bytp
Next bytR
If bytcontador = 0 Then
    SePuede = True
Else
    SePuede = False
End If
bytR = bytR + 1

End Function
Private Sub LlenarVec(ByRef vector() As Byte)
Dim bytR As Integer
Dim bytBi As Byte
Dim bytb As Byte
Dim bytK As Byte
Dim bytn As Byte

bytb = 1
bytR = 4
bytBi = 0

For bytn = 0 To 4
    For bytK = bytBi To bytBi + bytR
        vector(bytK) = bytb
    Next bytK
    bytR = bytR - 1
    bytBi = bytK
    bytb = bytb + 1
Next bytn

End Sub
Private Sub Swap(ByRef Var1 As Byte, ByRef Var2 As Byte)
Dim VarAux As Byte

VarAux = Var1
Var1 = Var2
Var2 = VarAux

End Sub
Private Sub DesordenarVec(ByRef vector() As Byte)
Dim byta As Byte
Dim bytb As Byte

    For byta = 0 To 14
        bytb = Fix(Rnd * 15)
        Swap vector(byta), vector(bytb)
    Next byta
    For byta = 0 To 14
        Print vector(byta)
    Next byta
End Sub
Private Sub Poner(ByVal fila As Integer, ByVal columna As Integer, ByRef matriz() As Byte, ByVal vector As Byte)
Dim bytp As Byte
Dim pos As Byte

For bytp = columna To columna + vector - 1
    matriz(fila, bytp) = vector
    pos = fila * 10 + bytp
    Agua(pos).Caption = matriz(fila, bytp)
    Agua(pos).BackColor = &H808000
Next bytp

End Sub
Private Sub btnEnd_Click()
End
End Sub

Private Sub btnGenerar_Click()
Dim matriz(9, 9) As Byte
Dim fila As Integer
Dim columna As Integer
Dim bytn As Byte
Dim contador As Byte
Dim vector(14) As Byte

LlenarVec vector()
DesordenarVec vector()

For bytn = 0 To 99
    Agua(bytn).BackColor = RGB(255, 255, 255)
    Agua(bytn).Caption = "0"
Next bytn

Do: DoEvents: contador = 0
    For bytn = 0 To 14
        fila = Fix(10 * Rnd): columna = Fix(Rnd * 10)
            If columna + vector(bytn) <= 10 Then
                If SePuede(fila, columna, matriz(), vector(bytn)) Then
                    Poner fila, columna, matriz(), vector(bytn)
                    contador = contador + 1
                End If
            End If
    Next bytn
Loop Until (contador = 15)
End Sub

Private Sub Form_Load()
Randomize Timer
End Sub
