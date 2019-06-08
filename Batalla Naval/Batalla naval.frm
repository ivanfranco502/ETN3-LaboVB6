VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnRecomenzar 
      Caption         =   "Borrar todo"
      Height          =   615
      Left            =   5280
      TabIndex        =   201
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton btnComenzar 
      Caption         =   "Generar  Barcos"
      Height          =   615
      Left            =   5280
      TabIndex        =   200
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   99
      Left            =   120
      TabIndex        =   199
      TabStop         =   0   'False
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   98
      Left            =   600
      TabIndex        =   198
      TabStop         =   0   'False
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   97
      Left            =   1080
      TabIndex        =   197
      TabStop         =   0   'False
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   96
      Left            =   1560
      TabIndex        =   196
      TabStop         =   0   'False
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   95
      Left            =   2040
      TabIndex        =   195
      TabStop         =   0   'False
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   94
      Left            =   2520
      TabIndex        =   194
      TabStop         =   0   'False
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   93
      Left            =   3000
      TabIndex        =   193
      TabStop         =   0   'False
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   92
      Left            =   3480
      TabIndex        =   192
      TabStop         =   0   'False
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   91
      Left            =   3960
      TabIndex        =   191
      TabStop         =   0   'False
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   90
      Left            =   4440
      TabIndex        =   190
      TabStop         =   0   'False
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   89
      Left            =   120
      TabIndex        =   189
      TabStop         =   0   'False
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   88
      Left            =   600
      TabIndex        =   188
      TabStop         =   0   'False
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   87
      Left            =   1080
      TabIndex        =   187
      TabStop         =   0   'False
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   86
      Left            =   1560
      TabIndex        =   186
      TabStop         =   0   'False
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   85
      Left            =   2040
      TabIndex        =   185
      TabStop         =   0   'False
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   84
      Left            =   2520
      TabIndex        =   184
      TabStop         =   0   'False
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   83
      Left            =   3000
      TabIndex        =   183
      TabStop         =   0   'False
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   82
      Left            =   3480
      TabIndex        =   182
      TabStop         =   0   'False
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   81
      Left            =   3960
      TabIndex        =   181
      TabStop         =   0   'False
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   80
      Left            =   4440
      TabIndex        =   180
      TabStop         =   0   'False
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   79
      Left            =   120
      TabIndex        =   179
      TabStop         =   0   'False
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   78
      Left            =   600
      TabIndex        =   178
      TabStop         =   0   'False
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   77
      Left            =   1080
      TabIndex        =   177
      TabStop         =   0   'False
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   76
      Left            =   1560
      TabIndex        =   176
      TabStop         =   0   'False
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   75
      Left            =   2040
      TabIndex        =   175
      TabStop         =   0   'False
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   74
      Left            =   2520
      TabIndex        =   174
      TabStop         =   0   'False
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   73
      Left            =   3000
      TabIndex        =   173
      TabStop         =   0   'False
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   72
      Left            =   3480
      TabIndex        =   172
      TabStop         =   0   'False
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   71
      Left            =   3960
      TabIndex        =   171
      TabStop         =   0   'False
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   70
      Left            =   4440
      TabIndex        =   170
      TabStop         =   0   'False
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   69
      Left            =   120
      TabIndex        =   169
      TabStop         =   0   'False
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   68
      Left            =   600
      TabIndex        =   168
      TabStop         =   0   'False
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   67
      Left            =   1080
      TabIndex        =   167
      TabStop         =   0   'False
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   66
      Left            =   1560
      TabIndex        =   166
      TabStop         =   0   'False
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   65
      Left            =   2040
      TabIndex        =   165
      TabStop         =   0   'False
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   64
      Left            =   2520
      TabIndex        =   164
      TabStop         =   0   'False
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   63
      Left            =   3000
      TabIndex        =   163
      TabStop         =   0   'False
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   62
      Left            =   3480
      TabIndex        =   162
      TabStop         =   0   'False
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   61
      Left            =   3960
      TabIndex        =   161
      TabStop         =   0   'False
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   60
      Left            =   4440
      TabIndex        =   160
      TabStop         =   0   'False
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   59
      Left            =   120
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   58
      Left            =   600
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   57
      Left            =   1080
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   56
      Left            =   1560
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   55
      Left            =   2040
      TabIndex        =   155
      TabStop         =   0   'False
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   54
      Left            =   2520
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   53
      Left            =   3000
      TabIndex        =   153
      TabStop         =   0   'False
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   52
      Left            =   3480
      TabIndex        =   152
      TabStop         =   0   'False
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   51
      Left            =   3960
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   50
      Left            =   4440
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   49
      Left            =   120
      TabIndex        =   149
      TabStop         =   0   'False
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   48
      Left            =   600
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   47
      Left            =   1080
      TabIndex        =   147
      TabStop         =   0   'False
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   46
      Left            =   1560
      TabIndex        =   146
      TabStop         =   0   'False
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   45
      Left            =   2040
      TabIndex        =   145
      TabStop         =   0   'False
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   44
      Left            =   2520
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   43
      Left            =   3000
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   42
      Left            =   3480
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   41
      Left            =   3960
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   40
      Left            =   4440
      TabIndex        =   140
      TabStop         =   0   'False
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   39
      Left            =   120
      TabIndex        =   139
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   38
      Left            =   600
      TabIndex        =   138
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   37
      Left            =   1080
      TabIndex        =   137
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   36
      Left            =   1560
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   35
      Left            =   2040
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   34
      Left            =   2520
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   33
      Left            =   3000
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   32
      Left            =   3480
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   31
      Left            =   3960
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   30
      Left            =   4440
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   29
      Left            =   120
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   28
      Left            =   600
      TabIndex        =   128
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   27
      Left            =   1080
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   26
      Left            =   1560
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   25
      Left            =   2040
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   24
      Left            =   2520
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   23
      Left            =   3000
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   22
      Left            =   3480
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   21
      Left            =   3960
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   20
      Left            =   4440
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   19
      Left            =   120
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   18
      Left            =   600
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   17
      Left            =   1080
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   16
      Left            =   1560
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   15
      Left            =   2040
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   14
      Left            =   2520
      TabIndex        =   114
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   13
      Left            =   3000
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   12
      Left            =   3480
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   11
      Left            =   3960
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   10
      Left            =   4440
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   9
      Left            =   4440
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   8
      Left            =   3960
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   7
      Left            =   3480
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   6
      Left            =   3000
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   5
      Left            =   2520
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   4
      Left            =   2040
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   1560
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   1080
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton btnCampo 
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   99
      Left            =   4440
      TabIndex        =   99
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   98
      Left            =   3960
      TabIndex        =   98
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   97
      Left            =   3480
      TabIndex        =   97
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   96
      Left            =   3000
      TabIndex        =   96
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   95
      Left            =   2520
      TabIndex        =   95
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   94
      Left            =   2040
      TabIndex        =   94
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   93
      Left            =   1560
      TabIndex        =   93
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   92
      Left            =   1080
      TabIndex        =   92
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   91
      Left            =   600
      TabIndex        =   91
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   90
      Left            =   120
      TabIndex        =   90
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   89
      Left            =   4440
      TabIndex        =   89
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   88
      Left            =   3960
      TabIndex        =   88
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   87
      Left            =   3480
      TabIndex        =   87
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   86
      Left            =   3000
      TabIndex        =   86
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   85
      Left            =   2520
      TabIndex        =   85
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   84
      Left            =   2040
      TabIndex        =   84
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   83
      Left            =   1560
      TabIndex        =   83
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   82
      Left            =   1080
      TabIndex        =   82
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   81
      Left            =   600
      TabIndex        =   81
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   80
      Left            =   120
      TabIndex        =   80
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   79
      Left            =   4440
      TabIndex        =   79
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   78
      Left            =   3960
      TabIndex        =   78
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   77
      Left            =   3480
      TabIndex        =   77
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   76
      Left            =   3000
      TabIndex        =   76
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   75
      Left            =   2520
      TabIndex        =   75
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   74
      Left            =   2040
      TabIndex        =   74
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   73
      Left            =   1560
      TabIndex        =   73
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   72
      Left            =   1080
      TabIndex        =   72
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   71
      Left            =   600
      TabIndex        =   71
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   70
      Left            =   120
      TabIndex        =   70
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   69
      Left            =   4440
      TabIndex        =   69
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   68
      Left            =   3960
      TabIndex        =   68
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   67
      Left            =   3480
      TabIndex        =   67
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   66
      Left            =   3000
      TabIndex        =   66
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   65
      Left            =   2520
      TabIndex        =   65
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   64
      Left            =   2040
      TabIndex        =   64
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   63
      Left            =   1560
      TabIndex        =   63
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   62
      Left            =   1080
      TabIndex        =   62
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   61
      Left            =   600
      TabIndex        =   61
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   60
      Left            =   120
      TabIndex        =   60
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   59
      Left            =   4440
      TabIndex        =   59
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   58
      Left            =   3960
      TabIndex        =   58
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   57
      Left            =   3480
      TabIndex        =   57
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   56
      Left            =   3000
      TabIndex        =   56
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   55
      Left            =   2520
      TabIndex        =   55
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   54
      Left            =   2040
      TabIndex        =   54
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   53
      Left            =   1560
      TabIndex        =   53
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   52
      Left            =   1080
      TabIndex        =   52
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   51
      Left            =   600
      TabIndex        =   51
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   50
      Left            =   120
      TabIndex        =   50
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   49
      Left            =   4440
      TabIndex        =   49
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   48
      Left            =   3960
      TabIndex        =   48
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   47
      Left            =   3480
      TabIndex        =   47
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   46
      Left            =   3000
      TabIndex        =   46
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   45
      Left            =   2520
      TabIndex        =   45
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   44
      Left            =   2040
      TabIndex        =   44
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   43
      Left            =   1560
      TabIndex        =   43
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   42
      Left            =   1080
      TabIndex        =   42
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   41
      Left            =   600
      TabIndex        =   41
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   40
      Left            =   120
      TabIndex        =   40
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   39
      Left            =   4440
      TabIndex        =   39
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   38
      Left            =   3960
      TabIndex        =   38
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   37
      Left            =   3480
      TabIndex        =   37
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   36
      Left            =   3000
      TabIndex        =   36
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   35
      Left            =   2520
      TabIndex        =   35
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   34
      Left            =   2040
      TabIndex        =   34
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   33
      Left            =   1560
      TabIndex        =   33
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   32
      Left            =   1080
      TabIndex        =   32
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   31
      Left            =   600
      TabIndex        =   31
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   30
      Left            =   120
      TabIndex        =   30
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   29
      Left            =   4440
      TabIndex        =   29
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   28
      Left            =   3960
      TabIndex        =   28
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   27
      Left            =   3480
      TabIndex        =   27
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   26
      Left            =   3000
      TabIndex        =   26
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   25
      Left            =   2520
      TabIndex        =   25
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   24
      Left            =   2040
      TabIndex        =   24
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   23
      Left            =   1560
      TabIndex        =   23
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   22
      Left            =   1080
      TabIndex        =   22
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   21
      Left            =   600
      TabIndex        =   21
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   20
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   19
      Left            =   4440
      TabIndex        =   19
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   18
      Left            =   3960
      TabIndex        =   18
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   17
      Left            =   3480
      TabIndex        =   17
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   16
      Left            =   3000
      TabIndex        =   16
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   15
      Left            =   2520
      TabIndex        =   15
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   14
      Left            =   2040
      TabIndex        =   14
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   13
      Left            =   1560
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   12
      Left            =   1080
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   11
      Left            =   600
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   9
      Left            =   4440
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   8
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   7
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   6
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   5
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   4
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCampo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
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
Option Explicit
Private Function diagonal(ByRef valor As Byte) As Boolean
Dim corroborar As Boolean
Select Case valor
    Case Is = 0
        If lblCampo(valor + 11).BackColor = &HFF& Then
            corroborar = False
        Else
            corroborar = True
        End If
    Case Is = 1, 2, 3, 4, 5, 6, 7, 8
        If lblCampo(valor + 11).BackColor = &HFF& Then
            corroborar = False
        Else
            If lblCampo(valor + 9).BackColor = &HFF& Then
                corroborar = False
            Else
                corroborar = True
            End If
        End If
    Case Is = 9
        If lblCampo(valor + 9).BackColor = &HFF& Then
            corroborar = False
        Else
            corroborar = True
        End If
    Case Is = 10, 20, 30, 40, 50, 60, 70, 80
        If lblCampo(valor - 9).BackColor = &HFF& Then
            corroborar = False
        Else
            If lblCampo(valor + 11).BackColor = &HFF& Then
                corroborar = False
            Else
                corroborar = True
            End If
        End If
    Case Is = 19, 29, 39, 49, 59, 69, 79, 89
        If lblCampo(valor - 11).BackColor = &HFF& Then
            corroborar = False
        Else
            If lblCampo(valor + 9).BackColor = &HFF& Then
                corroborar = False
            Else
                corroborar = True
            End If
        End If
    Case Is = 90
        If lblCampo(valor - 11).BackColor = &HFF& Then
            corroborar = False
        Else
            corroborar = True
        End If
    Case Is = 91, 92, 93, 94, 95, 96, 97, 98
        If lblCampo(valor - 11).BackColor = &HFF& Then
            corroborar = False
        Else
            If lblCampo(valor - 9).BackColor = &HFF& Then
                corroborar = False
            Else
                corroborar = True
            End If
        End If
    Case Is = 99
        If lblCampo(valor - 9).BackColor = &HFF& Then
            corroborar = False
        Else
            corroborar = True
        End If
    Case Else
        If lblCampo(valor - 11).BackColor = &HFF& Then
            corroborar = False
        Else
            If lblCampo(valor - 9).BackColor = &HFF& Then
                corroborar = False
            Else
                If lblCampo(valor + 11).BackColor = &HFF& Then
                    corroborar = False
                Else
                    If lblCampo(valor + 9).BackColor = &HFF& Then
                        corroborar = False
                    Else
                        corroborar = True
                    End If
                End If
            End If
        End If
End Select
diagonal = corroborar
End Function
Private Function Arriba(ByRef valor As Byte) As Boolean
Dim corroborar As Boolean
Select Case valor
    Case Is = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
        If lblCampo(valor + 10).BackColor = &HFF& Then
            corroborar = False
        Else
            corroborar = True
        End If
    Case Is = 90, 91, 92, 93, 94, 95, 96, 97, 98, 99
        If lblCampo(valor - 10).BackColor = &HFF& Then
            corroborar = False
        Else
            corroborar = True
        End If
    Case Else
        If lblCampo(valor - 10).BackColor = &HFF& Then
                corroborar = False
        Else
            If lblCampo(valor + 10).BackColor = &HFF& Then
                corroborar = False
            Else
                corroborar = True
            End If
        End If
Arriba = corroborar
End Select
End Function
Private Function Allado(ByRef valor As Byte) As Boolean
Dim corroborar As Boolean

Select Case valor
    Case Is = 9, 19, 29, 39, 49, 59, 69, 79, 89, 99
        If lblCampo(valor - 1).BackColor = &HFF& Then
            corroborar = False
        Else
            corroborar = True
        End If
    Case Is = 0, 10, 20, 30, 40, 50, 60, 70, 80, 90
        If lblCampo(valor + 1).BackColor = &HFF& Then
            corroborar = False
        Else
            corroborar = True
        End If
    Case Else
        If lblCampo(valor - 1).BackColor = &HFF& Then
            corroborar = False
        Else
            If lblCampo(valor + 1).BackColor = &HFF& Then
                corroborar = False
            Else
                corroborar = True
            End If
        End If
End Select
Allado = corroborar
End Function
Private Function igual(ByRef valor As Byte) As Boolean
If lblCampo(valor).BackColor = &HFF& Then
    igual = False
Else
    igual = True
End If
End Function
Private Sub llenado(bytvalor As Byte)
Static contador As Byte
Dim byta As Byte
Dim bytb As Byte
Select Case contador
    Case Is = 0
        For byta = 1 To 4
            If (igual(bytvalor) = True And (Arriba(bytvalor + byta) = True) And (diagonal(bytvalor + byta) = True)) Then
                If byta = 4 Then
                    If lblCampo(bytvalor + byta + 1).BackColor = &HFF& Then
                        If lblCampo(bytvalor - 2).BackColor = &HFF& Then
                            Print "no"
                        Else
                            lblCampo(bytvalor - 1).BackColor = &HFF&
                        End If
                    End If
                End If
            lblCampo(bytvalor + byta).BackColor = &HFF&
            End If
        Next byta
    Case Is = 1
        For bytb = 1 To 2
        For byta = 1 To 3
            If (igual(bytvalor) = True And (Arriba(bytvalor + byta) = True) And (diagonal(bytvalor + byta) = True)) Then
                If byta = 4 Then
                    If lblCampo(bytvalor + byta + 1).BackColor = &HFF& Then
                        If lblCampo(bytvalor - 2).BackColor = &HFF& Then
                            Print "no1"
                        Else
                            lblCampo(bytvalor - 1).BackColor = &HFF&
                        End If
                    End If
                End If
            lblCampo(bytvalor + byta).BackColor = &HFF&
            End If
        Next byta
        Next bytb
    Case Is = 2
        For bytb = 1 To 3
        For byta = 1 To 2
            If (igual(bytvalor) = True And (Arriba(bytvalor + byta) = True) And (diagonal(bytvalor + byta) = True)) Then
                If byta = 4 Then
                    If lblCampo(bytvalor + byta + 1).BackColor = &HFF& Then
                        If lblCampo(bytvalor - 2).BackColor = &HFF& Then
                            Print "no2"
                        Else
                            lblCampo(bytvalor - 1).BackColor = &HFF&
                        End If
                    End If
                End If
            lblCampo(bytvalor + byta).BackColor = &HFF&
            End If
        Next byta
        Next bytb
    Case Is = 3
        For bytb = 1 To 4
        For byta = 1 To 1
            If (igual(bytvalor) = True And (Arriba(bytvalor + byta) = True) And (diagonal(bytvalor + byta) = True)) Then
                If byta = 4 Then
                    If lblCampo(bytvalor + byta + 1).BackColor = &HFF& Then
                        If lblCampo(bytvalor - 2).BackColor = &HFF& Then
                            Print "no3"
                        Else
                            lblCampo(bytvalor - 1).BackColor = &HFF&
                        End If
                    End If
                End If
            lblCampo(bytvalor + byta).BackColor = &HFF&
            End If
        Next byta
        Next bytb
    Case Else
        contador = 0
End Select
contador = contador + 1
End Sub
Private Sub btnCampo_Click(Index As Integer)
btnCampo(Index).Visible = False
End Sub

Private Sub btnComenzar_Click()
Dim bytn As Byte
Dim byta As Byte
Dim bytvalor As Byte
Dim bytcontador As Byte

btnComenzar.Enabled = False
For bytn = 0 To 99
    btnCampo(bytn).Visible = False
Next bytn
byta = 0

Do: DoEvents
    Do: DoEvents
        bytvalor = Rnd * 99
    Loop Until (igual(bytvalor) = True And (Allado(bytvalor) = True) And (Arriba(bytvalor) = True) And (diagonal(bytvalor) = True))
    llenado (bytvalor)
    byta = byta + 1
    lblCampo(bytvalor).BackColor = &HFF&

Loop Until (byta = 15)

End Sub

Private Sub btnRecomenzar_Click()
Dim bytn As Byte
For bytn = 0 To 99
    btnCampo(bytn).Visible = True
    btnCampo(bytn).Enabled = False
    lblCampo(bytn).BackColor = &HFFFF80
Next bytn
btnComenzar.Enabled = True
End Sub

Private Sub Form_Load()
'1 de 5
'2 de 4
'3 de 3
'4 de 2
'5 de 1
Randomize (Timer)
End Sub


