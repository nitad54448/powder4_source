VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmControlFile 
   AutoRedraw      =   -1  'True
   Caption         =   "Set Rietveld Control File (DBWS 98)"
   ClientHeight    =   6710
   ClientLeft      =   50
   ClientTop       =   510
   ClientWidth     =   9790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6710
   ScaleWidth      =   9790
   Begin VB.Frame SetControlFrameTab 
      BorderStyle     =   0  'None
      Caption         =   "general"
      Height          =   5892
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   9252
      Begin VB.ListBox lstProfileFunction 
         Height          =   1040
         Left            =   5040
         TabIndex        =   223
         ToolTipText     =   "profile function"
         Top             =   4200
         Width           =   3012
      End
      Begin VB.ListBox lstDataFormat 
         Height          =   1040
         Left            =   5040
         TabIndex        =   222
         ToolTipText     =   "data file format"
         Top             =   1320
         Width           =   3012
      End
      Begin VB.CheckBox chkVariableInt 
         Caption         =   "variable intensity instrument"
         Height          =   492
         Left            =   480
         TabIndex        =   16
         Top             =   2880
         Width           =   2772
      End
      Begin VB.TextBox txtNrPhases 
         Height          =   288
         Left            =   2160
         TabIndex        =   6
         Text            =   "1"
         Top             =   4200
         Width           =   492
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Neutrons"
         Height          =   372
         Index           =   2
         Left            =   480
         TabIndex        =   5
         Top             =   2280
         Width           =   1812
      End
      Begin VB.CheckBox Check2 
         Caption         =   "X-rays"
         Height          =   372
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1812
      End
      Begin VB.CheckBox chkJobType 
         Caption         =   "Pattern calculation only"
         Height          =   252
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   2892
      End
      Begin VB.TextBox Text4 
         Height          =   288
         Left            =   480
         MaxLength       =   70
         TabIndex        =   2
         Text            =   "powder v3-gDBWS interface"
         Top             =   360
         Width           =   7572
      End
      Begin VB.Line lineSetControl 
         Index           =   1
         X1              =   360
         X2              =   8160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblSetControl 
         AutoSize        =   -1  'True
         Caption         =   "Job title"
         Height          =   192
         Index           =   5
         Left            =   360
         TabIndex        =   8
         Top             =   120
         Width           =   912
      End
      Begin VB.Line lineSetControl 
         Index           =   0
         X1              =   360
         X2              =   8160
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Number of phases"
         Height          =   312
         Index           =   6
         Left            =   360
         TabIndex        =   7
         Top             =   4200
         Width           =   1680
      End
   End
   Begin VB.Frame SetControlFrameTab 
      BorderStyle     =   0  'None
      Caption         =   "phase 1"
      Height          =   5892
      Index           =   7
      Left            =   240
      TabIndex        =   128
      Top             =   600
      Visible         =   0   'False
      Width           =   9252
      Begin VB.Frame Frame6 
         Caption         =   "atom no. 1"
         Height          =   3252
         Left            =   4680
         TabIndex        =   254
         Top             =   2520
         Width           =   4452
         Begin VB.TextBox atTxt 
            Height          =   288
            Index           =   2
            Left            =   1200
            TabIndex        =   262
            Top             =   1440
            Width           =   492
         End
         Begin VB.TextBox atTxt 
            Height          =   288
            Index           =   1
            Left            =   1200
            TabIndex        =   257
            Top             =   1080
            Width           =   492
         End
         Begin VB.TextBox atTxt 
            Height          =   288
            Index           =   0
            Left            =   1200
            TabIndex        =   256
            Top             =   720
            Width           =   492
         End
         Begin VB.HScrollBar atScroll 
            Height          =   252
            Left            =   1200
            Max             =   100
            Min             =   1
            TabIndex        =   255
            Top             =   0
            Value           =   1
            Width           =   372
         End
         Begin MSFlexGridLib.MSFlexGrid Grd 
            Height          =   2772
            Index           =   2
            Left            =   1920
            TabIndex        =   258
            Top             =   360
            Width           =   2412
            _ExtentX        =   4251
            _ExtentY        =   4886
            _Version        =   393216
            Rows            =   25
            Cols            =   3
            RowHeightMin    =   20
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483634
            GridColorFixed  =   -2147483630
            HighLight       =   0
            GridLines       =   2
            ScrollBars      =   2
            BorderStyle     =   0
            Appearance      =   0
         End
         Begin VB.Label lblatom2 
            Alignment       =   2  'Center
            Caption         =   "NTYP"
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   261
            Top             =   1440
            Width           =   732
         End
         Begin VB.Label lblatom2 
            Alignment       =   2  'Center
            Caption         =   "Mult."
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   260
            Top             =   1080
            Width           =   732
         End
         Begin VB.Label lblatom2 
            Alignment       =   2  'Center
            Caption         =   "LABEL"
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   259
            Top             =   720
            Width           =   732
         End
      End
      Begin VB.HScrollBar hscPhase 
         CausesValidation=   0   'False
         Height          =   240
         Left            =   120
         Max             =   15
         Min             =   1
         TabIndex        =   219
         Top             =   0
         Value           =   1
         Width           =   1212
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   0
         Left            =   4800
         TabIndex        =   208
         Top             =   0
         Width           =   4212
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   1
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   207
         Top             =   480
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   2
         Left            =   8400
         TabIndex        =   206
         Top             =   480
         Width           =   650
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   3
         Left            =   1560
         TabIndex        =   205
         Text            =   "1"
         Top             =   1680
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   4
         Left            =   1560
         TabIndex        =   204
         Top             =   2040
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   5
         Left            =   2160
         TabIndex        =   203
         Text            =   "0.001"
         Top             =   2040
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   6
         Left            =   2760
         TabIndex        =   202
         Top             =   2040
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   7
         Left            =   7080
         TabIndex        =   201
         Top             =   480
         Width           =   650
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   8
         Left            =   1560
         TabIndex        =   200
         Top             =   840
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   9
         Left            =   2160
         TabIndex        =   199
         Top             =   840
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   10
         Left            =   1560
         TabIndex        =   198
         Top             =   1200
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   11
         Left            =   2160
         TabIndex        =   197
         Top             =   1200
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   38
         Left            =   1560
         TabIndex        =   196
         Text            =   "1.0"
         Top             =   2400
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   39
         Left            =   1560
         TabIndex        =   195
         Top             =   2760
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   40
         Left            =   1560
         TabIndex        =   194
         Text            =   "0.16"
         Top             =   3120
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   41
         Left            =   2160
         TabIndex        =   193
         Top             =   2400
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   42
         Left            =   2160
         TabIndex        =   192
         Top             =   2760
         Width           =   500
      End
      Begin VB.TextBox txtPhase 
         Height          =   288
         Index           =   43
         Left            =   2160
         TabIndex        =   191
         Top             =   3120
         Width           =   500
      End
      Begin VB.Frame Frame4 
         Caption         =   "lattice parameters and codewords"
         Height          =   1572
         Left            =   4680
         TabIndex        =   172
         Top             =   840
         Width           =   4452
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   58
            Left            =   3720
            TabIndex        =   184
            Top             =   960
            Width           =   650
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   59
            Left            =   3000
            TabIndex        =   183
            Top             =   960
            Width           =   650
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   60
            Left            =   2280
            TabIndex        =   182
            Top             =   960
            Width           =   650
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   61
            Left            =   1560
            TabIndex        =   181
            Top             =   960
            Width           =   650
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   62
            Left            =   840
            TabIndex        =   180
            Top             =   960
            Width           =   650
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   63
            Left            =   120
            TabIndex        =   179
            Top             =   960
            Width           =   650
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   64
            Left            =   3720
            TabIndex        =   178
            Top             =   600
            Width           =   650
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   65
            Left            =   3000
            TabIndex        =   177
            Top             =   600
            Width           =   650
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   66
            Left            =   2280
            TabIndex        =   176
            Top             =   600
            Width           =   650
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   67
            Left            =   1560
            TabIndex        =   175
            Top             =   600
            Width           =   650
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   68
            Left            =   840
            TabIndex        =   174
            Top             =   600
            Width           =   650
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   69
            Left            =   120
            TabIndex        =   173
            Top             =   600
            Width           =   650
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "a"
            Height          =   192
            Index           =   74
            Left            =   120
            TabIndex        =   190
            Top             =   360
            Width           =   648
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "b"
            Height          =   192
            Index           =   75
            Left            =   840
            TabIndex        =   189
            Top             =   360
            Width           =   648
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "c"
            Height          =   192
            Index           =   76
            Left            =   1560
            TabIndex        =   188
            Top             =   360
            Width           =   648
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "alpha"
            Height          =   192
            Index           =   77
            Left            =   2280
            TabIndex        =   187
            Top             =   360
            Width           =   648
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "beta"
            Height          =   192
            Index           =   78
            Left            =   3000
            TabIndex        =   186
            Top             =   360
            Width           =   648
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "gamma"
            Height          =   192
            Index           =   79
            Left            =   3720
            TabIndex        =   185
            Top             =   360
            Width           =   648
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "profile parameters and codewords"
         Height          =   2172
         Left            =   120
         TabIndex        =   129
         Top             =   3600
         Width           =   4452
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   57
            Left            =   3840
            TabIndex        =   157
            Top             =   1800
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   56
            Left            =   3240
            TabIndex        =   156
            Top             =   1800
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   55
            Left            =   2640
            TabIndex        =   155
            Top             =   1800
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   54
            Left            =   2040
            TabIndex        =   154
            Top             =   1800
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   53
            Left            =   1440
            TabIndex        =   153
            Top             =   1800
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   52
            Left            =   840
            TabIndex        =   152
            Top             =   1800
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   51
            Left            =   240
            TabIndex        =   151
            Top             =   1800
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   50
            Left            =   3840
            TabIndex        =   150
            Top             =   1440
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   49
            Left            =   3240
            TabIndex        =   149
            Top             =   1440
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   48
            Left            =   2640
            TabIndex        =   148
            Top             =   1440
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   47
            Left            =   2040
            TabIndex        =   147
            Top             =   1440
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   46
            Left            =   1440
            TabIndex        =   146
            Top             =   1440
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   45
            Left            =   840
            TabIndex        =   145
            Top             =   1440
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   44
            Left            =   240
            TabIndex        =   144
            Text            =   "0.51"
            Top             =   1440
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   37
            Left            =   3840
            TabIndex        =   143
            Top             =   840
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   36
            Left            =   3240
            TabIndex        =   142
            Top             =   840
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   35
            Left            =   2640
            TabIndex        =   141
            Top             =   840
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   34
            Left            =   2040
            TabIndex        =   140
            Top             =   840
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   33
            Left            =   1440
            TabIndex        =   139
            Top             =   840
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   32
            Left            =   840
            TabIndex        =   138
            Top             =   840
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   31
            Left            =   240
            TabIndex        =   137
            Top             =   840
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   30
            Left            =   3840
            TabIndex        =   136
            Top             =   480
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   29
            Left            =   3240
            TabIndex        =   135
            Top             =   480
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   28
            Left            =   2640
            TabIndex        =   134
            Top             =   480
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   27
            Left            =   2040
            TabIndex        =   133
            Top             =   480
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   26
            Left            =   1440
            TabIndex        =   132
            Text            =   ".0059"
            Top             =   480
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   25
            Left            =   840
            TabIndex        =   131
            Text            =   "-0.0032"
            Top             =   480
            Width           =   500
         End
         Begin VB.TextBox txtPhase 
            Height          =   288
            Index           =   24
            Left            =   240
            TabIndex        =   130
            Text            =   "0.019"
            Top             =   480
            Width           =   500
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "U"
            Height          =   192
            Index           =   44
            Left            =   240
            TabIndex        =   171
            Top             =   240
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "X"
            Height          =   192
            Index           =   59
            Left            =   3240
            TabIndex        =   170
            Top             =   240
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "Z"
            Height          =   192
            Index           =   60
            Left            =   2640
            TabIndex        =   169
            Top             =   240
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "CT"
            Height          =   192
            Index           =   61
            Left            =   2040
            TabIndex        =   168
            Top             =   240
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "W"
            Height          =   192
            Index           =   62
            Left            =   1440
            TabIndex        =   167
            Top             =   240
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "V"
            Height          =   192
            Index           =   63
            Left            =   840
            TabIndex        =   166
            Top             =   240
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "Y"
            Height          =   192
            Index           =   64
            Left            =   3840
            TabIndex        =   165
            Top             =   240
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "NA"
            Height          =   192
            Index           =   67
            Left            =   240
            TabIndex        =   164
            Top             =   1200
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "NB"
            Height          =   192
            Index           =   68
            Left            =   840
            TabIndex        =   163
            Top             =   1200
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "NC"
            Height          =   192
            Index           =   69
            Left            =   1440
            TabIndex        =   162
            Top             =   1200
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "hNA"
            Height          =   192
            Index           =   70
            Left            =   2040
            TabIndex        =   161
            Top             =   1200
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "hNB"
            Height          =   192
            Index           =   71
            Left            =   2640
            TabIndex        =   160
            Top             =   1200
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "hNC"
            Height          =   192
            Index           =   72
            Left            =   3240
            TabIndex        =   159
            Top             =   1200
            Width           =   504
         End
         Begin VB.Label lblSetControl 
            Alignment       =   2  'Center
            Caption         =   "A"
            Height          =   192
            Index           =   73
            Left            =   3840
            TabIndex        =   158
            Top             =   1200
            Width           =   504
         End
      End
      Begin VB.Label lblPhaseSet 
         Alignment       =   2  'Center
         Caption         =   "PHASE 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   1440
         TabIndex        =   220
         Top             =   0
         Width           =   1212
      End
      Begin VB.Label lblSetControl 
         Alignment       =   2  'Center
         Caption         =   "phase name"
         Height          =   312
         Index           =   35
         Left            =   3480
         TabIndex        =   127
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "no. of atoms"
         Height          =   312
         Index           =   36
         Left            =   240
         TabIndex        =   218
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "Z"
         Height          =   312
         Index           =   37
         Left            =   7680
         TabIndex        =   217
         Top             =   480
         Width           =   576
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "Tp, abs. factor"
         Height          =   312
         Index           =   38
         Left            =   120
         TabIndex        =   216
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "pref. orientation"
         Height          =   312
         Index           =   39
         Left            =   120
         TabIndex        =   215
         Top             =   2040
         Width           =   1332
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "space group"
         Height          =   312
         Index           =   40
         Left            =   5760
         TabIndex        =   214
         Top             =   480
         Width           =   1152
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "scale factor"
         Height          =   312
         Index           =   42
         Left            =   240
         TabIndex        =   213
         Top             =   840
         Width           =   1188
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "overall thermal parameter"
         Height          =   432
         Index           =   43
         Left            =   120
         TabIndex        =   212
         Top             =   1200
         Width           =   1332
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "G1"
         Height          =   192
         Index           =   45
         Left            =   876
         TabIndex        =   211
         Top             =   2400
         Width           =   504
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "G2"
         Height          =   192
         Index           =   65
         Left            =   876
         TabIndex        =   210
         Top             =   2760
         Width           =   504
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "P"
         Height          =   192
         Index           =   66
         Left            =   840
         TabIndex        =   209
         Top             =   3120
         Width           =   504
      End
   End
   Begin VB.Frame SetControlFrameTab 
      BorderStyle     =   0  'None
      Caption         =   "Refinable global parameters"
      Height          =   5772
      Index           =   6
      Left            =   240
      TabIndex        =   73
      Top             =   600
      Visible         =   0   'False
      Width           =   9012
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   31
         Left            =   7080
         TabIndex        =   123
         Top             =   4800
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   30
         Left            =   6000
         TabIndex        =   122
         Top             =   4800
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   29
         Left            =   4920
         TabIndex        =   121
         Top             =   4800
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   28
         Left            =   3840
         TabIndex        =   120
         Top             =   4800
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   27
         Left            =   2760
         TabIndex        =   119
         Top             =   4800
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   26
         Left            =   1680
         TabIndex        =   118
         Top             =   4800
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   25
         Left            =   7080
         TabIndex        =   117
         Top             =   4440
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   24
         Left            =   6000
         TabIndex        =   116
         Top             =   4440
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   23
         Left            =   4920
         TabIndex        =   115
         Top             =   4440
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   22
         Left            =   3840
         TabIndex        =   114
         Top             =   4440
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   21
         Left            =   2760
         TabIndex        =   113
         Top             =   4440
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   19
         Left            =   7200
         TabIndex        =   102
         Top             =   3240
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   18
         Left            =   2880
         TabIndex        =   101
         Top             =   3240
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   17
         Left            =   2880
         TabIndex        =   100
         Top             =   2520
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   9
         Left            =   2880
         TabIndex        =   99
         Top             =   1800
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   13
         Left            =   7200
         TabIndex        =   98
         Top             =   2520
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   12
         Left            =   7200
         TabIndex        =   97
         Top             =   1800
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   11
         Left            =   7200
         TabIndex        =   96
         Top             =   1080
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   10
         Left            =   7200
         TabIndex        =   95
         Top             =   360
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   8
         Left            =   2880
         TabIndex        =   94
         Top             =   1080
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   7
         Left            =   2880
         TabIndex        =   93
         Top             =   360
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   20
         Left            =   1680
         TabIndex        =   92
         Top             =   4440
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   16
         Left            =   6360
         TabIndex        =   91
         Top             =   3240
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   15
         Left            =   2040
         TabIndex        =   90
         Top             =   3240
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   14
         Left            =   2040
         TabIndex        =   82
         Top             =   2520
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   2
         Left            =   2040
         TabIndex        =   81
         Top             =   1800
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   6
         Left            =   6360
         TabIndex        =   80
         Top             =   2520
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   5
         Left            =   6360
         TabIndex        =   79
         Top             =   1800
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   4
         Left            =   6360
         TabIndex        =   78
         Top             =   1080
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   3
         Left            =   6360
         TabIndex        =   77
         Top             =   360
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   1
         Left            =   2040
         TabIndex        =   76
         Top             =   1080
         Width           =   732
      End
      Begin VB.TextBox txtGlobalRefine 
         Height          =   288
         Index           =   0
         Left            =   2040
         TabIndex        =   74
         Top             =   360
         Width           =   732
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   8040
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label lblSetControl 
         Alignment       =   2  'Center
         Caption         =   "codewords"
         Height          =   312
         Index           =   51
         Left            =   7080
         TabIndex        =   126
         Top             =   0
         Width           =   972
      End
      Begin VB.Label lblSetControl 
         Alignment       =   2  'Center
         Caption         =   "codewords"
         Height          =   312
         Index           =   50
         Left            =   2760
         TabIndex        =   125
         Top             =   0
         Width           =   972
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "codewords"
         Height          =   312
         Index           =   34
         Left            =   240
         TabIndex        =   124
         Top             =   4800
         Width           =   1236
      End
      Begin VB.Label lblSetControl 
         Alignment       =   2  'Center
         Caption         =   "background coeff. for 5th polynomial "
         Height          =   312
         Index           =   49
         Left            =   240
         TabIndex        =   112
         Top             =   4080
         Width           =   3132
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "MON2"
         Height          =   312
         Index           =   48
         Left            =   4800
         TabIndex        =   111
         Top             =   3240
         Width           =   1476
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "MON1"
         Height          =   312
         Index           =   47
         Left            =   360
         TabIndex        =   110
         Top             =   3240
         Width           =   1572
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "scale factor for Tape 11"
         Height          =   312
         Index           =   33
         Left            =   120
         TabIndex        =   89
         Top             =   2520
         Width           =   1812
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "surface roughness, T"
         Height          =   312
         Index           =   32
         Left            =   4200
         TabIndex        =   88
         Top             =   2520
         Width           =   2100
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "surface roughness, R"
         Height          =   312
         Index           =   31
         Left            =   4080
         TabIndex        =   87
         Top             =   1800
         Width           =   2232
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "surface roughness, Q"
         Height          =   312
         Index           =   30
         Left            =   4200
         TabIndex        =   86
         Top             =   1080
         Width           =   2112
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "surface roughness, P"
         Height          =   252
         Index           =   29
         Left            =   4200
         TabIndex        =   85
         Top             =   360
         Width           =   2100
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "transparency"
         Height          =   312
         Index           =   28
         Left            =   360
         TabIndex        =   84
         Top             =   1800
         Width           =   1536
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "sample displacement"
         Height          =   312
         Index           =   27
         Left            =   120
         TabIndex        =   83
         Top             =   1080
         Width           =   1788
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         Caption         =   "offset of 2theta-zero"
         Height          =   312
         Index           =   26
         Left            =   120
         TabIndex        =   75
         Top             =   360
         Width           =   1860
      End
   End
   Begin VB.Frame SetControlFrameTab 
      BorderStyle     =   0  'None
      Caption         =   "Run choices"
      Height          =   5412
      Index           =   5
      Left            =   600
      TabIndex        =   60
      Top             =   720
      Width           =   8892
      Begin VB.TextBox txtParam 
         Height          =   288
         Left            =   6720
         TabIndex        =   267
         Top             =   240
         Width           =   612
      End
      Begin VB.Frame frameRunChoices 
         Caption         =   "pattern calculation"
         Enabled         =   0   'False
         Height          =   3852
         Left            =   4200
         TabIndex        =   103
         Top             =   1080
         Width           =   3972
         Begin VB.TextBox txtOperations 
            Height          =   288
            Index           =   6
            Left            =   2520
            TabIndex        =   109
            Top             =   840
            Width           =   732
         End
         Begin VB.TextBox txtOperations 
            Height          =   288
            Index           =   7
            Left            =   2520
            TabIndex        =   107
            Top             =   1680
            Width           =   732
         End
         Begin VB.TextBox txtOperations 
            Height          =   288
            Index           =   8
            Left            =   2520
            TabIndex        =   105
            Top             =   2520
            Width           =   732
         End
         Begin VB.Label lblSetControl 
            Alignment       =   1  'Right Justify
            Caption         =   "starting angle, 2theta"
            Height          =   372
            Index           =   23
            Left            =   240
            TabIndex        =   108
            Top             =   840
            Width           =   2172
         End
         Begin VB.Label lblSetControl 
            Alignment       =   1  'Right Justify
            Caption         =   "ending angle, deg. 2theta"
            Height          =   312
            Index           =   24
            Left            =   120
            TabIndex        =   106
            Top             =   1680
            Width           =   2268
         End
         Begin VB.Label lblSetControl 
            Alignment       =   1  'Right Justify
            Caption         =   "step size, deg. 2 theta"
            Height          =   312
            Index           =   25
            Left            =   360
            TabIndex        =   104
            Top             =   2520
            Width           =   2016
         End
      End
      Begin VB.TextBox txtOperations 
         Height          =   288
         Index           =   5
         Left            =   3000
         TabIndex        =   66
         Text            =   "0.95"
         Top             =   4440
         Width           =   612
      End
      Begin VB.TextBox txtOperations 
         Height          =   288
         Index           =   4
         Left            =   3000
         TabIndex        =   65
         Text            =   "0.95"
         Top             =   3600
         Width           =   612
      End
      Begin VB.TextBox txtOperations 
         Height          =   288
         Index           =   3
         Left            =   3000
         TabIndex        =   64
         Text            =   "0.95"
         Top             =   2760
         Width           =   612
      End
      Begin VB.TextBox txtOperations 
         Height          =   288
         Index           =   2
         Left            =   3000
         TabIndex        =   63
         Text            =   "0.95"
         Top             =   1920
         Width           =   612
      End
      Begin VB.TextBox txtOperations 
         Height          =   288
         Index           =   1
         Left            =   3000
         TabIndex        =   62
         Text            =   "0.1"
         Top             =   1080
         Width           =   612
      End
      Begin VB.TextBox txtOperations 
         Height          =   288
         Index           =   0
         Left            =   1920
         TabIndex        =   61
         Text            =   "3"
         Top             =   240
         Width           =   612
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "number of refined param."
         Height          =   192
         Index           =   1
         Left            =   4836
         TabIndex        =   268
         Top             =   240
         Width           =   1776
      End
      Begin VB.Label lblSetControl 
         Alignment       =   2  'Center
         Caption         =   "RELAX 4 (2theta zero, specimen displacement, transparency, roughness, amorphous)"
         Height          =   672
         Index           =   22
         Left            =   120
         TabIndex        =   72
         Top             =   4320
         Width           =   2772
      End
      Begin VB.Label lblSetControl 
         Alignment       =   2  'Center
         Caption         =   "RELAX 3 (profile width, asymmetry, overall atom displacement, preferred orientation, lattice, scale factor)"
         Height          =   792
         Index           =   21
         Left            =   240
         TabIndex        =   71
         Top             =   3480
         Width           =   2640
      End
      Begin VB.Label lblSetControl 
         Alignment       =   2  'Center
         Caption         =   "RELAX 2 (aniso atom displacement factors)"
         Height          =   672
         Index           =   20
         Left            =   240
         TabIndex        =   70
         Top             =   2640
         Width           =   2604
      End
      Begin VB.Label lblSetControl 
         Alignment       =   2  'Center
         Caption         =   "RELAX 1 (coordinates, isotropic atomic displacement factors, site occupancies)"
         Height          =   672
         Index           =   19
         Left            =   240
         TabIndex        =   69
         Top             =   1800
         Width           =   2712
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "stop when shifts reach limit, EPS "
         Height          =   312
         Index           =   18
         Left            =   360
         TabIndex        =   68
         Top             =   1080
         Width           =   2532
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "number of cycles"
         Height          =   312
         Index           =   17
         Left            =   240
         TabIndex        =   67
         Top             =   240
         Width           =   1572
      End
   End
   Begin VB.Frame SetControlFrameTab 
      BorderStyle     =   0  'None
      Caption         =   "experiment"
      Height          =   5292
      Index           =   4
      Left            =   480
      TabIndex        =   41
      Top             =   840
      Width           =   8652
      Begin VB.TextBox txtFixedPar 
         Height          =   288
         Index           =   8
         Left            =   7320
         TabIndex        =   50
         Text            =   "0.1"
         Top             =   3360
         Width           =   972
      End
      Begin VB.TextBox txtFixedPar 
         Height          =   288
         Index           =   7
         Left            =   7320
         TabIndex        =   49
         Text            =   "40"
         Top             =   2520
         Width           =   972
      End
      Begin VB.TextBox txtFixedPar 
         Height          =   288
         Index           =   6
         Left            =   7320
         TabIndex        =   48
         Text            =   "1"
         Top             =   1560
         Width           =   972
      End
      Begin VB.TextBox txtFixedPar 
         Height          =   288
         Index           =   5
         Left            =   7320
         TabIndex        =   47
         Text            =   "1"
         Top             =   600
         Width           =   972
      End
      Begin VB.TextBox txtFixedPar 
         Height          =   288
         Index           =   4
         Left            =   2760
         TabIndex        =   46
         Text            =   "8"
         Top             =   3360
         Width           =   972
      End
      Begin VB.TextBox txtFixedPar 
         Height          =   288
         Index           =   3
         Left            =   2760
         TabIndex        =   45
         Text            =   "90"
         Top             =   2520
         Width           =   972
      End
      Begin VB.TextBox txtFixedPar 
         Height          =   288
         Index           =   2
         Left            =   2760
         TabIndex        =   44
         Text            =   "0.5"
         Top             =   1560
         Width           =   972
      End
      Begin VB.TextBox txtFixedPar 
         Height          =   288
         Index           =   1
         Left            =   2760
         TabIndex        =   43
         Text            =   " 1.54439"
         Top             =   1080
         Width           =   972
      End
      Begin VB.TextBox txtFixedPar 
         Height          =   288
         Index           =   0
         Left            =   2760
         TabIndex        =   42
         Text            =   " 1.540562"
         Top             =   600
         Width           =   972
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "sample thickness, (cm)"
         Height          =   312
         Index           =   16
         Left            =   4800
         TabIndex        =   59
         Top             =   3360
         Width           =   2352
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "asymmetry correction limit, 2theta"
         Height          =   312
         Index           =   15
         Left            =   4560
         TabIndex        =   58
         Top             =   2520
         Width           =   2592
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "linear abs. coefficient, (1/cm)"
         Height          =   432
         Index           =   14
         Left            =   4920
         TabIndex        =   57
         Top             =   1560
         Width           =   2280
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "monochromator coefficient, LP"
         Height          =   432
         Index           =   13
         Left            =   4320
         TabIndex        =   56
         Top             =   600
         Width           =   2868
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "width of calculated profile"
         Height          =   312
         Index           =   12
         Left            =   240
         TabIndex        =   55
         Top             =   3360
         Width           =   2388
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "background, origin of polynomial"
         Height          =   312
         Index           =   11
         Left            =   120
         TabIndex        =   54
         Top             =   2520
         Width           =   2532
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "intensity ratio, w2/w1"
         Height          =   312
         Index           =   10
         Left            =   240
         TabIndex        =   53
         Top             =   1560
         Width           =   2412
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "wavelength 2"
         Height          =   312
         Index           =   9
         Left            =   840
         TabIndex        =   52
         Top             =   1080
         Width           =   1776
      End
      Begin VB.Label lblSetControl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "wavelength 1 "
         Height          =   312
         Index           =   8
         Left            =   600
         TabIndex        =   51
         Top             =   600
         Width           =   2052
      End
   End
   Begin VB.Frame SetControlFrameTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "output flags"
      Height          =   5652
      Index           =   3
      Left            =   720
      TabIndex        =   19
      Top             =   480
      Width           =   8532
      Begin VB.Frame Frame3 
         Caption         =   "Riello background"
         Height          =   4692
         Left            =   4560
         TabIndex        =   33
         Top             =   600
         Width           =   2892
         Begin VB.CheckBox Check6 
            Caption         =   "Iobs coorected (ALL)"
            Height          =   492
            Index           =   6
            Left            =   240
            TabIndex        =   40
            Top             =   3480
            Value           =   1  'Checked
            Width           =   2292
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Amorphous "
            Height          =   492
            Index           =   5
            Left            =   240
            TabIndex        =   39
            Top             =   3000
            Width           =   1692
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Disorder"
            Height          =   492
            Index           =   4
            Left            =   240
            TabIndex        =   38
            Top             =   2520
            Width           =   1692
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Compton scattering"
            Height          =   492
            Index           =   3
            Left            =   240
            TabIndex        =   37
            Top             =   1920
            Width           =   1692
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Polyomial background"
            Height          =   492
            Index           =   2
            Left            =   240
            TabIndex        =   36
            Top             =   1440
            Width           =   2292
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Icalc"
            Height          =   492
            Index           =   1
            Left            =   240
            TabIndex        =   35
            Top             =   840
            Width           =   1692
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Io corrected for absorption"
            Height          =   492
            Index           =   0
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   2532
         End
      End
      Begin VB.CheckBox Check5 
         Caption         =   "last cycle parameters and esd's"
         Height          =   372
         Index           =   12
         Left            =   480
         TabIndex        =   32
         Top             =   4920
         Width           =   3492
      End
      Begin VB.CheckBox Check5 
         Caption         =   "summary at each cycle"
         Height          =   372
         Index           =   11
         Left            =   480
         TabIndex        =   31
         Top             =   4560
         Value           =   1  'Checked
         Width           =   2772
      End
      Begin VB.CheckBox Check5 
         Caption         =   "symmetry operators"
         Height          =   372
         Index           =   10
         Left            =   480
         TabIndex        =   30
         Top             =   4200
         Width           =   3612
      End
      Begin VB.CheckBox Check5 
         Caption         =   "merged reflection list"
         Height          =   372
         Index           =   9
         Left            =   480
         TabIndex        =   29
         Top             =   3840
         Width           =   2892
      End
      Begin VB.CheckBox Check5 
         Caption         =   "corrected data list"
         Height          =   372
         Index           =   8
         Left            =   480
         TabIndex        =   28
         Top             =   3480
         Width           =   2172
      End
      Begin VB.CheckBox Check5 
         Caption         =   "reflection list"
         Height          =   372
         Index           =   7
         Left            =   480
         TabIndex        =   27
         Top             =   3120
         Width           =   1452
      End
      Begin VB.CheckBox Check5 
         Caption         =   "update input file"
         Height          =   372
         Index           =   6
         Left            =   480
         TabIndex        =   26
         Top             =   2760
         Width           =   1452
      End
      Begin VB.CheckBox Check5 
         Caption         =   "correlation matrix"
         Height          =   372
         Index           =   5
         Left            =   480
         TabIndex        =   25
         Top             =   2400
         Width           =   1812
      End
      Begin VB.CheckBox Check5 
         Caption         =   "str.factors, R-B, R-F, A+iB"
         Height          =   372
         Index           =   4
         Left            =   480
         TabIndex        =   24
         Top             =   2040
         Value           =   1  'Checked
         Width           =   3252
      End
      Begin VB.CheckBox Check5 
         Caption         =   "str.factors, R-B, R-F, Fo, Fc, phase angles"
         Height          =   372
         Index           =   3
         Left            =   480
         TabIndex        =   23
         Top             =   1680
         Width           =   3492
      End
      Begin VB.CheckBox Check5 
         Caption         =   "str. factors and R-Bragg"
         Height          =   372
         Index           =   2
         Left            =   480
         TabIndex        =   22
         Top             =   1320
         Width           =   2532
      End
      Begin VB.CheckBox Check5 
         Caption         =   "line printer plot file"
         Height          =   372
         Index           =   1
         Left            =   480
         TabIndex        =   21
         Top             =   960
         Width           =   3012
      End
      Begin VB.CheckBox Check5 
         Caption         =   "obs. and calc. intensities at each step"
         Height          =   372
         Index           =   0
         Left            =   480
         TabIndex        =   20
         Top             =   600
         Width           =   3732
      End
   End
   Begin VB.Frame SetControlFrameTab 
      BorderStyle     =   0  'None
      Caption         =   "other parameters, global"
      Height          =   5772
      Index           =   2
      Left            =   480
      TabIndex        =   15
      Top             =   600
      Width           =   8772
      Begin VB.Frame frm 
         Caption         =   "Excluded regions"
         Height          =   1452
         Index           =   4
         Left            =   120
         TabIndex        =   238
         Top             =   4200
         Width           =   5172
         Begin VB.TextBox txtExcl 
            Height          =   288
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   241
            Text            =   "0"
            Top             =   480
            Width           =   612
         End
         Begin MSFlexGridLib.MSFlexGrid Grd 
            Height          =   852
            Index           =   0
            Left            =   2160
            TabIndex        =   239
            Top             =   480
            Width           =   2772
            _ExtentX        =   4886
            _ExtentY        =   1499
            _Version        =   393216
            FixedCols       =   0
            RowHeightMin    =   20
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483634
            GridColorFixed  =   -2147483630
            Enabled         =   -1  'True
            HighLight       =   0
            GridLines       =   2
            ScrollBars      =   2
            BorderStyle     =   0
            Appearance      =   0
         End
         Begin VB.Label Label1 
            Caption         =   "how many ?"
            Height          =   612
            Index           =   4
            Left            =   120
            TabIndex        =   240
            Top             =   480
            Width           =   852
         End
      End
      Begin VB.Frame frm 
         Caption         =   "Scattering factors (X-rays)"
         Height          =   3852
         Index           =   3
         Left            =   120
         TabIndex        =   235
         Top             =   120
         Width           =   5172
         Begin VB.HScrollBar hsc2 
            Enabled         =   0   'False
            Height          =   252
            Left            =   2160
            Max             =   21
            Min             =   1
            TabIndex        =   253
            Top             =   1200
            Value           =   1
            Width           =   732
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   2172
            Left            =   240
            TabIndex        =   242
            Top             =   1560
            Width           =   1692
            Begin VB.TextBox txtScat 
               Enabled         =   0   'False
               Height          =   288
               Index           =   3
               Left            =   960
               TabIndex        =   250
               Top             =   1680
               Width           =   612
            End
            Begin VB.TextBox txtScat 
               Enabled         =   0   'False
               Height          =   288
               Index           =   2
               Left            =   960
               TabIndex        =   249
               Top             =   1200
               Width           =   612
            End
            Begin VB.TextBox txtScat 
               Enabled         =   0   'False
               Height          =   288
               Index           =   1
               Left            =   960
               TabIndex        =   248
               Top             =   720
               Width           =   612
            End
            Begin VB.TextBox txtScat 
               Enabled         =   0   'False
               Height          =   288
               Index           =   0
               Left            =   960
               TabIndex        =   247
               Top             =   240
               Width           =   612
            End
            Begin VB.Label lbl2 
               Caption         =   "at. weight"
               Height          =   252
               Index           =   3
               Left            =   0
               TabIndex        =   246
               Top             =   1680
               Width           =   852
            End
            Begin VB.Label lbl2 
               Caption         =   "Im. part"
               Height          =   252
               Index           =   2
               Left            =   0
               TabIndex        =   245
               Top             =   1200
               Width           =   852
            End
            Begin VB.Label lbl2 
               Caption         =   "Re. part"
               Height          =   252
               Index           =   1
               Left            =   0
               TabIndex        =   244
               Top             =   720
               Width           =   852
            End
            Begin VB.Label lbl2 
               Caption         =   "Name"
               Height          =   372
               Index           =   0
               Left            =   0
               TabIndex        =   243
               Top             =   240
               Width           =   852
            End
         End
         Begin VB.CheckBox chkIntTable 
            Caption         =   "Int'l Tables"
            Height          =   252
            Left            =   3600
            TabIndex        =   237
            Top             =   480
            Width           =   1092
         End
         Begin VB.CheckBox chkScatt 
            Caption         =   "add scattering factors"
            Height          =   492
            Left            =   240
            TabIndex        =   236
            Top             =   360
            Width           =   3372
         End
         Begin MSFlexGridLib.MSFlexGrid Grd 
            Height          =   1692
            Index           =   1
            Left            =   2160
            TabIndex        =   251
            Top             =   1800
            Width           =   2772
            _ExtentX        =   4886
            _ExtentY        =   2981
            _Version        =   393216
            Rows            =   10
            Cols            =   3
            FixedRows       =   0
            RowHeightMin    =   20
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483634
            GridColorFixed  =   -2147483630
            ScrollTrack     =   -1  'True
            Enabled         =   0   'False
            HighLight       =   0
            GridLines       =   2
            ScrollBars      =   2
            AllowUserResizing=   3
            BorderStyle     =   0
            Appearance      =   0
         End
         Begin VB.Label lblscatt 
            Caption         =   "set no:"
            Height          =   252
            Left            =   360
            TabIndex        =   252
            Top             =   1200
            Width           =   1212
         End
      End
      Begin VB.Frame frm 
         Caption         =   "Surface roughness"
         Height          =   2052
         Index           =   2
         Left            =   5640
         TabIndex        =   230
         Top             =   3600
         Width           =   3132
         Begin VB.OptionButton Option3 
            Caption         =   "combination model"
            Height          =   252
            Index           =   0
            Left            =   360
            TabIndex        =   234
            Top             =   480
            Value           =   -1  'True
            Width           =   2052
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Sparks et al."
            Height          =   252
            Index           =   1
            Left            =   360
            TabIndex        =   233
            Top             =   840
            Width           =   2052
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Suortti  "
            Height          =   252
            Index           =   2
            Left            =   360
            TabIndex        =   232
            Top             =   1200
            Width           =   2052
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Pitschke et al."
            Height          =   252
            Index           =   3
            Left            =   360
            TabIndex        =   231
            Top             =   1560
            Width           =   2052
         End
      End
      Begin VB.Frame frm 
         Caption         =   "Preffered orientation"
         Height          =   1332
         Index           =   1
         Left            =   5640
         TabIndex        =   227
         Top             =   1920
         Width           =   3132
         Begin VB.OptionButton Option1 
            Caption         =   "March-Dollase function"
            Height          =   372
            Index           =   1
            Left            =   360
            TabIndex        =   229
            Top             =   720
            Value           =   -1  'True
            Width           =   2292
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rietveld-Toraya function"
            Height          =   372
            Index           =   0
            Left            =   360
            TabIndex        =   228
            Top             =   360
            Width           =   2292
         End
      End
      Begin VB.Frame frm 
         Caption         =   "Asymmetry"
         Height          =   1332
         Index           =   0
         Left            =   5640
         TabIndex        =   224
         Top             =   120
         Width           =   3132
         Begin VB.OptionButton Option2 
            Caption         =   "Riello et al. model"
            Height          =   372
            Index           =   1
            Left            =   360
            TabIndex        =   226
            Top             =   720
            Width           =   1932
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Rietveld model"
            Height          =   372
            Index           =   0
            Left            =   360
            TabIndex        =   225
            Top             =   360
            Value           =   -1  'True
            Width           =   2412
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "surface roughness"
         Height          =   1932
         Left            =   5280
         TabIndex        =   17
         Top             =   3120
         Width           =   2892
      End
   End
   Begin VB.Frame SetControlFrameTab 
      BorderStyle     =   0  'None
      Caption         =   " 1"
      Height          =   5532
      Index           =   1
      Left            =   600
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   8532
      Begin VB.Frame Frame7 
         Caption         =   "alternate background"
         Enabled         =   0   'False
         Height          =   2172
         Left            =   4560
         TabIndex        =   263
         Top             =   2760
         Width           =   3612
         Begin VB.OptionButton OPTb 
            Caption         =   "Overall temp. factors"
            Enabled         =   0   'False
            Height          =   372
            Index           =   2
            Left            =   360
            TabIndex        =   266
            Top             =   1560
            Width           =   3012
         End
         Begin VB.OptionButton OPTb 
            Caption         =   "Individual temp. factors"
            Enabled         =   0   'False
            Height          =   372
            Index           =   1
            Left            =   360
            TabIndex        =   265
            Top             =   960
            Width           =   3132
         End
         Begin VB.OptionButton OPTb 
            Caption         =   "Standard background"
            Enabled         =   0   'False
            Height          =   372
            Index           =   0
            Left            =   360
            TabIndex        =   264
            Top             =   360
            Width           =   3012
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "use absorbtion correction (Riello)"
         Enabled         =   0   'False
         Height          =   372
         Index           =   0
         Left            =   4920
         TabIndex        =   18
         Top             =   2160
         Width           =   3012
      End
      Begin VB.OptionButton optSetControlBackground 
         Caption         =   "Riello et al."
         Height          =   372
         Index           =   3
         Left            =   4920
         TabIndex        =   13
         Top             =   1440
         Width           =   1452
      End
      Begin VB.OptionButton optSetControlBackground 
         Caption         =   "linear interpolation, points:"
         Height          =   372
         Index           =   2
         Left            =   480
         TabIndex        =   12
         Top             =   2160
         Width           =   2412
      End
      Begin VB.OptionButton optSetControlBackground 
         Caption         =   "read background from file, tape 3"
         Height          =   372
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Width           =   3012
      End
      Begin VB.OptionButton optSetControlBackground 
         Caption         =   "5th order polynomial "
         Height          =   372
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Value           =   -1  'True
         Width           =   2652
      End
      Begin MSFlexGridLib.MSFlexGrid Grd 
         Height          =   2172
         Index           =   3
         Left            =   480
         TabIndex        =   221
         Top             =   2760
         Width           =   2892
         _ExtentX        =   5098
         _ExtentY        =   3828
         _Version        =   393216
         Rows            =   30
         Cols            =   3
         RowHeightMin    =   20
         BackColorBkg    =   -2147483633
         GridColor       =   -2147483634
         GridColorFixed  =   -2147483630
         Enabled         =   0   'False
         HighLight       =   0
         GridLines       =   2
         ScrollBars      =   2
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.Line lineSetControl 
         Index           =   3
         X1              =   360
         X2              =   3000
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lblSetControl 
         AutoSize        =   -1  'True
         Caption         =   "Background model"
         Height          =   192
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   120
         Width           =   1368
      End
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   240
      Top             =   6360
      _ExtentX        =   670
      _ExtentY        =   670
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin ComctlLib.TabStrip tabControlFile 
      Height          =   6492
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9612
      _ExtentX        =   16951
      _ExtentY        =   11448
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   8
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Background"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "other parameters, global"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "output flags"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "experiment"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "run choices"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "refinable global"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "phase information"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mainCF 
      Caption         =   "Control File"
      Begin VB.Menu mnuSetControlFile 
         Caption         =   "Save control file (Pw_icf.txt)"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCFOpen 
         Caption         =   "Open Control File"
      End
      Begin VB.Menu mnuSaveControlFile 
         Caption         =   "Save Control File"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpty 
         Caption         =   "Clear Control File"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCFPrint 
         Caption         =   "Print Conditions"
         Visible         =   0   'False
      End
      Begin VB.Menu s3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCFExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mCFRunDBWS 
      Caption         =   "gDBWS v3"
      Begin VB.Menu mnuCFRun 
         Caption         =   "Run"
      End
      Begin VB.Menu lm1 
         Caption         =   "-"
      End
      Begin VB.Menu mRunSelectData 
         Caption         =   "Select data File"
         Begin VB.Menu mSelectOnDisk 
            Caption         =   "on Disk"
         End
         Begin VB.Menu mDataInMemory 
            Caption         =   "in Memory"
         End
      End
      Begin VB.Menu lm2 
         Caption         =   "-"
      End
      Begin VB.Menu mUpgradeCF 
         Caption         =   "Upgrade Control File"
         Checked         =   -1  'True
      End
      Begin VB.Menu mShowResults 
         Caption         =   "Show Results File"
         Checked         =   -1  'True
      End
      Begin VB.Menu mShowCFGraph 
         Caption         =   "Show plotinfo File"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuCFImportatoms 
      Caption         =   "Import atoms"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmControlFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentScatteringSet As Integer
Dim ScattSetCallFromOpen As Boolean

Private Sub atScroll_Change()
On Error GoTo errortrap
If CInt(CStr(txtPhase(1).Text)) > 200 Then MsgBox "The number of atoms is restricted to 200 (already  too many...)": txtPhase(1).Text = "200": Exit Sub
atScroll.Max = CInt(CStr(txtPhase(1).Text))
'update the atomdata here (take the last value of atscroll.value)
'if everything is right go on
'temp is a variant
''MsgBox "phaseedited" & CStr(phaseEdited)
Call refreshAtomData
Frame6.Caption = "atom no. " & (CStr(atScroll.Value))
AtomEdited = atScroll.Value
Call FormRefreshAtomData
Exit Sub
errortrap:
RietveldBoardMessage "error encountered in atScroll routine"
Exit Sub
End Sub

Sub SetScatteringFactors()
'read the scatt factors only if the check box is activated
Dim i As Integer
If chkScatt.Value = Unchecked Then totalScatt = 0: Exit Sub
' CurrentScatteringSet is an integer showing the last number for the set to be added
'it can be only up to 20
'Type ScatteringFactors
'IntTable As Boolean
'Name As String
'RePart As Single
'ImPart As Single
'AtWeight As Single
'NineCoeff(9) As Single
'PosScatt(2, 50) As Single
'End Type
'''Dim Scattering(20) As ScatteringFactors
Scattering(CurrentScatteringSet).IntTable = False
If chkIntTable.Value = Checked Then Scattering(CurrentScatteringSet).IntTable = True
Scattering(CurrentScatteringSet).Name = txtScat(0).Text
Scattering(CurrentScatteringSet).RePart = CSng(Val(txtScat(1).Text))
Scattering(CurrentScatteringSet).ImPart = CSng(Val(txtScat(2).Text))
Scattering(CurrentScatteringSet).AtWeight = CSng(Val(txtScat(3).Text))
If Scattering(CurrentScatteringSet).IntTable Then
'read the 9 coefficients in the string nineCoeff
For i = 1 To 9
Scattering(CurrentScatteringSet).NineCoeff(i) = CSng(Val(GridVal(Grd(1), i - 1, 1)))
Next i
Else
'read pos, scatt up to 50 sets of values
For i = 1 To 100
If i + 1 > Grd(1).Rows Then Exit For
Scattering(CurrentScatteringSet).PosScatt(1, i) = CSng(Val(GridVal(Grd(1), i - 1, 1)))
Scattering(CurrentScatteringSet).PosScatt(2, i) = CSng(Val(GridVal(Grd(1), i - 1, 2)))
Next i
End If
End Sub

Sub FormScatteringFactors()
'write the currentset of scattering factors
Dim i As Integer
If chkScatt.Value = Unchecked Then Exit Sub
'currentscatteringset is important
If Scattering(CurrentScatteringSet).IntTable Then
chkIntTable.Value = Checked
Else
chkIntTable.Value = Unchecked
End If
txtScat(0).Text = Scattering(CurrentScatteringSet).Name
txtScat(1).Text = CStr(Scattering(CurrentScatteringSet).RePart)
txtScat(2).Text = CStr(Scattering(CurrentScatteringSet).ImPart)
txtScat(3).Text = CStr(Scattering(CurrentScatteringSet).AtWeight)

If txtScat(1).Text = "0" Then txtScat(1).Text = ""
If txtScat(2).Text = "0" Then txtScat(2).Text = ""
If txtScat(3).Text = "0" Then txtScat(3).Text = ""

If Scattering(CurrentScatteringSet).IntTable Then
'put the 9 coefficients in the string nineCoeff
For i = 1 To 9: Call SetGridVal(Grd(1), i - 1, 1, CStr(Scattering(CurrentScatteringSet).NineCoeff(i))): Next i
Else
'read pos, scatt up to 50 sets of values

For i = 1 To 100
If i + 1 >= Grd(1).Rows Or (Val(Scattering(CurrentScatteringSet).PosScatt(2, i)) = 0) Then Exit For
Call SetGridVal(Grd(1), i - 1, 1, CStr(Scattering(CurrentScatteringSet).PosScatt(1, i)))
Call SetGridVal(Grd(1), i - 1, 2, CStr(Scattering(CurrentScatteringSet).PosScatt(2, i)))
Next i
End If

Exit Sub
End Sub

Sub FormRefreshAtomData()
'actualize all the atom data on the form,...
On Error GoTo errortrap
atTxt(0).Text = CStr(Atoms(AtomEdited, phaseEdited).Label)
atTxt(1).Text = CStr(Atoms(AtomEdited, phaseEdited).Multiplicity)
If Val(atTxt(1).Text) <= 0 Then atTxt(1).Text = ""
atTxt(2).Text = CStr(Atoms(AtomEdited, phaseEdited).Ntyp)
    
Call SetGridVal(Grd(2), 1, 1, CStr(Atoms(AtomEdited, phaseEdited).X))
Call SetGridVal(Grd(2), 2, 1, CStr(Atoms(AtomEdited, phaseEdited).Y))
Call SetGridVal(Grd(2), 3, 1, CStr(Atoms(AtomEdited, phaseEdited).z))
Call SetGridVal(Grd(2), 4, 1, CStr(Atoms(AtomEdited, phaseEdited).IsotropicThermal))
Call SetGridVal(Grd(2), 5, 1, CStr(Atoms(AtomEdited, phaseEdited).SiteOccupancy))
Call SetGridVal(Grd(2), 6, 1, CStr(Atoms(AtomEdited, phaseEdited).Beta11))
Call SetGridVal(Grd(2), 7, 1, CStr(Atoms(AtomEdited, phaseEdited).Beta22))
Call SetGridVal(Grd(2), 8, 1, CStr(Atoms(AtomEdited, phaseEdited).Beta33))
Call SetGridVal(Grd(2), 9, 1, CStr(Atoms(AtomEdited, phaseEdited).Beta12))
Call SetGridVal(Grd(2), 10, 1, CStr(Atoms(AtomEdited, phaseEdited).Beta13))
Call SetGridVal(Grd(2), 11, 1, CStr(Atoms(AtomEdited, phaseEdited).Beta23))

Call SetGridVal(Grd(2), 1, 2, CStr(Atoms(AtomEdited, phaseEdited).codeX))
Call SetGridVal(Grd(2), 2, 2, CStr(Atoms(AtomEdited, phaseEdited).codeY))
Call SetGridVal(Grd(2), 3, 2, CStr(Atoms(AtomEdited, phaseEdited).codeZ))
Call SetGridVal(Grd(2), 4, 2, CStr(Atoms(AtomEdited, phaseEdited).codeIsotropicThermal))
Call SetGridVal(Grd(2), 5, 2, CStr(Atoms(AtomEdited, phaseEdited).codeSiteOccupancy))
Call SetGridVal(Grd(2), 6, 2, CStr(Atoms(AtomEdited, phaseEdited).codeBeta11))
Call SetGridVal(Grd(2), 7, 2, CStr(Atoms(AtomEdited, phaseEdited).codeBeta22))
Call SetGridVal(Grd(2), 8, 2, CStr(Atoms(AtomEdited, phaseEdited).codeBeta33))
Call SetGridVal(Grd(2), 9, 2, CStr(Atoms(AtomEdited, phaseEdited).codeBeta12))
Call SetGridVal(Grd(2), 10, 2, CStr(Atoms(AtomEdited, phaseEdited).codeBeta13))
Call SetGridVal(Grd(2), 11, 2, CStr(Atoms(AtomEdited, phaseEdited).codeBeta23))
Exit Sub
errortrap:
RietveldBoardMessage strLinie
RietveldBoardMessage "error in formRefreshAtomData routine.."
Exit Sub












End Sub


Sub FormRefreshPhaseData()
'actualize on the form all the data related to the phase
On Error GoTo errortrap

txtPhase(0).Text = Phases(phaseEdited).Name
txtPhase(1).Text = CStr(Phases(phaseEdited).Atomi)
txtPhase(2).Text = CStr(Phases(phaseEdited).FormulaUnits)
txtPhase(3).Text = CStr(Phases(phaseEdited).ParticleAbsorptionFactor)
txtPhase(4).Text = CStr(Phases(phaseEdited).PrefOrientation(1))
txtPhase(5).Text = CStr(Phases(phaseEdited).PrefOrientation(2))
txtPhase(6).Text = CStr(Phases(phaseEdited).PrefOrientation(3))
txtPhase(7).Text = CStr(Phases(phaseEdited).SpaceGroupSymbol)
txtPhase(8).Text = CStr(Phases(phaseEdited).scalefactor)
txtPhase(9).Text = CStr(Phases(phaseEdited).scaleFactorCode)
txtPhase(10).Text = CStr(Phases(phaseEdited).OverallThermal)
txtPhase(11).Text = CStr(Phases(phaseEdited).OverallThermalCode)

For i = 1 To 11
If txtPhase(i).Text = "0" Then txtPhase(i).Text = ""
Next i

txtPhase(24).Text = CStr(Phases(phaseEdited).U)
txtPhase(25).Text = CStr(Phases(phaseEdited).v)
txtPhase(26).Text = CStr(Phases(phaseEdited).W)
txtPhase(27).Text = CStr(Phases(phaseEdited).CT)
txtPhase(28).Text = CStr(Phases(phaseEdited).z)
txtPhase(29).Text = CStr(Phases(phaseEdited).X)
txtPhase(30).Text = CStr(Phases(phaseEdited).Y)
txtPhase(31).Text = CStr(Phases(phaseEdited).codeU)
txtPhase(32).Text = CStr(Phases(phaseEdited).codeV)
txtPhase(33).Text = CStr(Phases(phaseEdited).codeW)
txtPhase(34).Text = CStr(Phases(phaseEdited).codeCT)
txtPhase(35).Text = CStr(Phases(phaseEdited).codeZ)
txtPhase(36).Text = CStr(Phases(phaseEdited).codeX)
txtPhase(37).Text = CStr(Phases(phaseEdited).codeY)
txtPhase(38).Text = CStr(Phases(phaseEdited).G1)
txtPhase(39).Text = CStr(Phases(phaseEdited).G2)
txtPhase(40).Text = CStr(Phases(phaseEdited).P)
txtPhase(41).Text = CStr(Phases(phaseEdited).codeG1)
txtPhase(42).Text = CStr(Phases(phaseEdited).codeG2)
txtPhase(43).Text = CStr(Phases(phaseEdited).codeP)
txtPhase(44).Text = CStr(Phases(phaseEdited).NA)
txtPhase(45).Text = CStr(Phases(phaseEdited).NB)
txtPhase(46).Text = CStr(Phases(phaseEdited).NC)
txtPhase(47).Text = CStr(Phases(phaseEdited).hNA)
txtPhase(48).Text = CStr(Phases(phaseEdited).hNB)
txtPhase(49).Text = CStr(Phases(phaseEdited).hNC)
txtPhase(50).Text = CStr(Phases(phaseEdited).SP7A)

txtPhase(51).Text = CStr(Phases(phaseEdited).codeNA)
txtPhase(52).Text = CStr(Phases(phaseEdited).codeNB)
txtPhase(53).Text = CStr(Phases(phaseEdited).codeNC)
txtPhase(54).Text = CStr(Phases(phaseEdited).codehNA)
txtPhase(55).Text = CStr(Phases(phaseEdited).codehNB)
txtPhase(56).Text = CStr(Phases(phaseEdited).codehNC)
txtPhase(57).Text = CStr(Phases(phaseEdited).codeSP7A)
txtPhase(58).Text = CStr(Phases(phaseEdited).codegamma)
txtPhase(59).Text = CStr(Phases(phaseEdited).codeBeta)
txtPhase(60).Text = CStr(Phases(phaseEdited).codeAlpha)
txtPhase(61).Text = CStr(Phases(phaseEdited).codeC)
txtPhase(62).Text = CStr(Phases(phaseEdited).codeB)
txtPhase(63).Text = CStr(Phases(phaseEdited).codeA)
txtPhase(64).Text = CStr(Phases(phaseEdited).gamma)
txtPhase(65).Text = CStr(Phases(phaseEdited).Beta)
txtPhase(66).Text = CStr(Phases(phaseEdited).Alpha)
txtPhase(67).Text = CStr(Phases(phaseEdited).c)
txtPhase(68).Text = CStr(Phases(phaseEdited).b)
txtPhase(69).Text = CStr(Phases(phaseEdited).a)
For i = 24 To 69
If txtPhase(i).Text = "0" Then txtPhase(i).Text = ""
Next i


Exit Sub
errortrap:

MsgBox "Error in actualizing the Phase data..please check the output file carefully."
Exit Sub
End Sub



Function GridVal(grila As Object, i As Integer, j As Integer) As String
'returns the value of the string which is in the object Grid, at pozition i,j
'Dim cod As Boolean, i As Integer, j As Integer
grila.Col = j: grila.Row = i
GridVal = grila.Text
If Val(GridVal) = 0 Then GridVal = ""
Exit Function
End Function

Sub SetGridVal(grila As Object, i As Integer, j As Integer, s As String)
'sets the value in the object grid of coordinates i and j
grila.Col = j: grila.Row = i
grila.Text = s
Exit Sub
End Sub


Sub refreshPhaseData()
On Error GoTo errortrap
Phases(phaseEdited).Atomi = CInt(txtPhase(1).Text)
Phases(phaseEdited).scalefactor = CSng(Val(txtPhase(8).Text))
Phases(phaseEdited).scaleFactorCode = CSng(Val(txtPhase(9).Text))
Phases(phaseEdited).OverallThermal = CSng(Val(txtPhase(10).Text))
Phases(phaseEdited).OverallThermalCode = CSng(Val(txtPhase(11).Text))
Phases(phaseEdited).ParticleAbsorptionFactor = CSng(Val(txtPhase(3).Text))
Phases(phaseEdited).PrefOrientation(1) = CSng(Val(txtPhase(4).Text))
Phases(phaseEdited).PrefOrientation(2) = CSng(Val(txtPhase(5).Text))
Phases(phaseEdited).PrefOrientation(3) = CSng(Val(txtPhase(6).Text))
Phases(phaseEdited).G1 = CSng(Val(txtPhase(38).Text))
Phases(phaseEdited).G2 = CSng(Val(txtPhase(39).Text))
Phases(phaseEdited).P = CSng(Val(txtPhase(40).Text))
Phases(phaseEdited).codeG1 = CSng(Val(txtPhase(41).Text))
Phases(phaseEdited).codeG2 = CSng(Val(txtPhase(42).Text))
Phases(phaseEdited).codeP = CSng(Val(txtPhase(43).Text))

Phases(phaseEdited).U = CSng(Val(txtPhase(24).Text))
Phases(phaseEdited).v = CSng(Val(txtPhase(25).Text))
Phases(phaseEdited).W = CSng(Val(txtPhase(26).Text))
Phases(phaseEdited).CT = CSng(Val(txtPhase(27).Text))
Phases(phaseEdited).z = CSng(Val(txtPhase(28).Text))
Phases(phaseEdited).X = CSng(Val(txtPhase(29).Text))
Phases(phaseEdited).Y = CSng(Val(txtPhase(30).Text))
Phases(phaseEdited).codeU = CSng(Val(txtPhase(31).Text))
Phases(phaseEdited).codeV = CSng(Val(txtPhase(32).Text))
Phases(phaseEdited).codeW = CSng(Val(txtPhase(33).Text))
Phases(phaseEdited).codeCT = CSng(Val(txtPhase(34).Text))
Phases(phaseEdited).codeZ = CSng(Val(txtPhase(35).Text))
Phases(phaseEdited).codeX = CSng(Val(txtPhase(36).Text))
Phases(phaseEdited).codeY = CSng(Val(txtPhase(37).Text))

Phases(phaseEdited).NA = CSng(Val(txtPhase(44).Text))
Phases(phaseEdited).NB = CSng(Val(txtPhase(45).Text))
Phases(phaseEdited).NC = CSng(Val(txtPhase(46).Text))
Phases(phaseEdited).hNA = CSng(Val(txtPhase(47).Text))
Phases(phaseEdited).hNB = CSng(Val(txtPhase(48).Text))
Phases(phaseEdited).hNC = CSng(Val(txtPhase(49).Text))
Phases(phaseEdited).SP7A = CSng(Val(txtPhase(50).Text))
Phases(phaseEdited).codeNA = CSng(Val(txtPhase(51).Text))
Phases(phaseEdited).codeNB = CSng(Val(txtPhase(52).Text))
Phases(phaseEdited).codeNC = CSng(Val(txtPhase(53).Text))
Phases(phaseEdited).codehNA = CSng(Val(txtPhase(54).Text))
Phases(phaseEdited).codehNB = CSng(Val(txtPhase(55).Text))
Phases(phaseEdited).codehNC = CSng(Val(txtPhase(56).Text))
Phases(phaseEdited).codeSP7A = CSng(Val(txtPhase(57).Text))

Phases(phaseEdited).Name = CStr(txtPhase(0).Text)
Phases(phaseEdited).SpaceGroupSymbol = CStr(txtPhase(7).Text)
Phases(phaseEdited).FormulaUnits = CInt(Val(txtPhase(2).Text))

Phases(phaseEdited).a = CSng(Val(txtPhase(69).Text))
Phases(phaseEdited).b = CSng(Val(txtPhase(68).Text))
Phases(phaseEdited).c = CSng(Val(txtPhase(67).Text))
Phases(phaseEdited).Alpha = CSng(Val(txtPhase(66).Text))
Phases(phaseEdited).Beta = CSng(Val(txtPhase(65).Text))
Phases(phaseEdited).gamma = CSng(Val(txtPhase(64).Text))

Phases(phaseEdited).codeA = CSng(Val(txtPhase(63).Text))
Phases(phaseEdited).codeB = CSng(Val(txtPhase(62).Text))
Phases(phaseEdited).codeC = CSng(Val(txtPhase(61).Text))
Phases(phaseEdited).codeAlpha = CSng(Val(txtPhase(60).Text))
Phases(phaseEdited).codeBeta = CSng(Val(txtPhase(59).Text))
Phases(phaseEdited).codegamma = CSng(Val(txtPhase(58).Text))
''RietveldBoardMessage "updating phase data...done"

Exit Sub
errortrap:
RietveldBoardMessage strLinie
RietveldBoardMessage "warning: error in updating phase data..."
RietveldBoardMessage Err.Description
Exit Sub
End Sub

Sub refreshAtomData()
On Error GoTo errortrap
 Atoms(AtomEdited, phaseEdited).Label = atTxt(0).Text
    Atoms(AtomEdited, phaseEdited).Multiplicity = Val(atTxt(1).Text)

    If Len(atTxt(2).Text) > 4 Then atTxt(2).Text = Mid$(atTxt(2).Text, 1, 4)
    Do Until Len(atTxt(2).Text) = 4
    atTxt(2).Text = atTxt(2).Text + " "
    Loop
    If Len(atTxt(2).Text) > 4 Then atTxt(2).Text = Mid$(atTxt(2).Text, 1, 4)
    Atoms(AtomEdited, phaseEdited).Ntyp = CStr(UCase$(atTxt(2).Text))
    'Ntyp it should be 4 char long
    Atoms(AtomEdited, phaseEdited).X = CSng(Val(GridVal(Grd(2), 1, 1)))
    Atoms(AtomEdited, phaseEdited).Y = CSng(Val(GridVal(Grd(2), 2, 1)))
    Atoms(AtomEdited, phaseEdited).z = CSng(Val(GridVal(Grd(2), 3, 1)))
    Atoms(AtomEdited, phaseEdited).IsotropicThermal = CSng(Val(GridVal(Grd(2), 4, 1)))
    Atoms(AtomEdited, phaseEdited).SiteOccupancy = CSng(Val(GridVal(Grd(2), 5, 1)))
    Atoms(AtomEdited, phaseEdited).Beta11 = CSng(Val(GridVal(Grd(2), 6, 1)))
    Atoms(AtomEdited, phaseEdited).Beta22 = CSng(Val(GridVal(Grd(2), 7, 1)))
    Atoms(AtomEdited, phaseEdited).Beta33 = CSng(Val(GridVal(Grd(2), 8, 1)))
    Atoms(AtomEdited, phaseEdited).Beta12 = CSng(Val(GridVal(Grd(2), 9, 1)))
    Atoms(AtomEdited, phaseEdited).Beta13 = CSng(Val(GridVal(Grd(2), 10, 1)))
    Atoms(AtomEdited, phaseEdited).Beta23 = CSng(Val(GridVal(Grd(2), 11, 1)))

    Atoms(AtomEdited, phaseEdited).codeX = CSng(Val(GridVal(Grd(2), 1, 2)))
    Atoms(AtomEdited, phaseEdited).codeY = CSng(Val(GridVal(Grd(2), 2, 2)))
    Atoms(AtomEdited, phaseEdited).codeZ = CSng(Val(GridVal(Grd(2), 3, 2)))
    Atoms(AtomEdited, phaseEdited).codeIsotropicThermal = CSng(Val(GridVal(Grd(2), 4, 2)))
    Atoms(AtomEdited, phaseEdited).codeSiteOccupancy = CSng(Val(GridVal(Grd(2), 5, 2)))
    Atoms(AtomEdited, phaseEdited).codeBeta11 = CSng(Val(GridVal(Grd(2), 6, 2)))
    Atoms(AtomEdited, phaseEdited).codeBeta22 = CSng(Val(GridVal(Grd(2), 7, 2)))
    Atoms(AtomEdited, phaseEdited).codeBeta33 = CSng(Val(GridVal(Grd(2), 8, 2)))
    Atoms(AtomEdited, phaseEdited).codeBeta12 = CSng(Val(GridVal(Grd(2), 9, 2)))
    Atoms(AtomEdited, phaseEdited).codeBeta13 = CSng(Val(GridVal(Grd(2), 10, 2)))
    Atoms(AtomEdited, phaseEdited).codeBeta23 = CSng(Val(GridVal(Grd(2), 11, 2)))
Exit Sub
errortrap:
RietveldBoardMessage strLinie
RietveldBoardMessage "error in RefreshAtomData routine.."
Exit Sub
End Sub


Private Sub Check2_Click(Index As Integer)

'set up the xrays or neutron experiment
Select Case Index

Case 1
If Check2(1).Value = 1 Then
Check2(2).Value = 0
Else
Check2(2).Value = 1
End If

Case 2
If Check2(2).Value = 1 Then
Check2(1).Value = 0
Else
Check2(1).Value = 1
End If
End Select

End Sub

Sub chkIntTable_Click()
'if this is On, the set-up of the grid will be made as for the international table
'int Tables ON
Select Case chkIntTable.Value
Case 1
Grd(1).Cols = 2
Grd(1).Rows = 9
Grd(1).Col = 0
Grd(1).ColWidth(0) = 2 * Grd(1).Width / 5.2
Grd(1).ColWidth(1) = 3 * Grd(1).Width / 5.4
Grd(1).Row = 0: Grd(1).Text = "A1"
Grd(1).Row = 1: Grd(1).Text = "B1"
Grd(1).Row = 2: Grd(1).Text = "A2"
Grd(1).Row = 3: Grd(1).Text = "B2"
Grd(1).Row = 4: Grd(1).Text = "A3"
Grd(1).Row = 5: Grd(1).Text = "B3"
Grd(1).Row = 6: Grd(1).Text = "A4"
Grd(1).Row = 7: Grd(1).Text = "B4"
Grd(1).Row = 8: Grd(1).Text = "C"
'Grd(1).Row = 9: Grd(1).Text = ""

'if the Int Tables is OFF, the default,...
Case 0
Grd(1).Rows = 20
Grd(1).Cols = 3
Grd(1).ColWidth(0) = 2 * Grd(1).Width / 7.4
Grd(1).ColWidth(1) = 2.5 * Grd(1).Width / 7.4
Grd(1).ColWidth(2) = 2.5 * Grd(1).Width / 7.4
Grd(1).Col = 0

For i = 0 To Grd(1).Rows - 1
Grd(1).Row = i: Grd(1).Text = "posi, scat,"
Next i
End Select

End Sub

Private Sub chkScatt_Click()
Dim s As String, i As Integer
If ScattSetCallFromOpen Then Exit Sub

Grd(1).Enabled = False
lblscatt.Enabled = False
chkScatt.Caption = "add scattering factors"
lblscatt.Caption = "set no:"
For i = 0 To 3: txtScat(i).Enabled = False: Next i
hsc2.Enabled = False
If chkScatt.Value = 1 Then
s = InputBox("How many scattering factors sets to add (max 20) ?", "scattering factors", 1)
If Val(s) > 0 And Val(s) < 21 Then
chkScatt.Caption = chkScatt.Caption & ", " & CStr(CInt(Val((s))))
lblscatt.Caption = lblscatt.Caption & " " & "1"
hsc2.Value = 1
Grd(1).Enabled = True
lblscatt.Enabled = True
hsc2.Enabled = True
hsc2.Max = CInt(Val((s)))
totalScatt = s
'the four texts
For i = 0 To 3: txtScat(i).Enabled = True: Next i
'which way to write these factors
MsgBox ("You can insert here only XRD scattering factors. For neutrons insert only Name and Atomic weight; in the file insert the scattering length." & " In the X-rays case, for each set you can input either nine coefficients as listed in Int'l Tables or sets of sin(theta)/lambda , scattering factor, (one pair per line) ")
Else
s = ""
chkScatt.Value = 0
totalScatt = 0
End If
    'make the grid for data input


End If
DoEvents
End Sub

Private Sub Form_Load()
Dim coderoare As Boolean
On Error GoTo errortrap
'read the number of phases from txtnrphases
AtomEdited = 1: phaseEdited = 1: CurrentScatteringSet = 1
Dim i As Integer
i = CInt(txtNrPhases.Text)
If i < 1 Then
txtNrPhases.Text = 1
Else
If i > 15 Then
txtNrPhases.Text = 15
End If
End If
'set up the listboxes for data type and profile

lstDataFormat.AddItem (" standard DBWS")
lstDataFormat.AddItem (" free format")
lstDataFormat.AddItem (" standard GSAS")
lstDataFormat.AddItem (" Philips UDF")
lstDataFormat.AddItem (" Scintag text")

lstDataFormat.ListIndex = 0

lstProfileFunction.AddItem " Gaussian"
lstProfileFunction.AddItem " Lorentzian"
lstProfileFunction.AddItem " mod 1 Lorentzian"
lstProfileFunction.AddItem " mod 2 Lorentzian"
lstProfileFunction.AddItem " split Pearson VII "
lstProfileFunction.AddItem " pseudo Voigt"
lstProfileFunction.AddItem " Pearson VII (symmetric)"
lstProfileFunction.AddItem " mod Thomson Cox Hastings"
lstProfileFunction.ListIndex = 4
'set the bar to the number of phases
hscPhase.Min = 1
hscPhase.Max = Val(txtNrPhases.Text)

'set up the grid sizes and stuff
'the four grids are: 3 for background, 2 for the atoms, 1 for scattering, 0 for the excluded regions
'put titles, set the sizes of the columns and rows, etc...
    'this is the excluded region, number
    Call SetGridVal(Grd(0), 0, 0, "start 2t/deg")
    Call SetGridVal(Grd(0), 0, 1, "end 2t/deg")
'set the width of the columns for the excluded region grid
    Grd(0).ColWidth(0) = 4.8 * Grd(0).Width / 10
    Grd(0).ColWidth(1) = 4.8 * Grd(0).Width / 10
    
'set the width of the columns for the scattering factors
'warning: there are two different ways of writting the scattering factors
    Grd(1).ColWidth(0) = Grd(1).Width / 3
    Grd(1).ColWidth(1) = Grd(1).Width / 3

'2 is for the atom parameter
    Grd(2).ColWidth(0) = 1.5 * Grd(2).Width / 6
    Grd(2).ColWidth(1) = 2 * Grd(2).Width / 6
    Grd(2).ColWidth(2) = 2 * Grd(2).Width / 6
'write here all necessary
Grd(2).Col = 1: Grd(2).Row = 0: Grd(2).Text = "value": Grd(2).Col = 2: Grd(2).Text = "code"
'I have to write here x, y, z B, So,beta 11, 22, 33, 12, 13, 23
Grd(2).Col = 0: Grd(2).Row = 1: Grd(2).Text = "X": Grd(2).Row = 2: Grd(2).Text = "Y": Grd(2).Row = 3: Grd(2).Text = "Z":
Grd(2).Row = 4: Grd(2).Text = "B": Grd(2).Row = 5: Grd(2).Text = "So": Grd(2).Row = 6: Grd(2).Text = "Beta11":
Grd(2).Row = 7: Grd(2).Text = "Beta22": Grd(2).Row = 8: Grd(2).Text = "Beta33": Grd(2).Row = 9: Grd(2).Text = "Beta12": Grd(2).Row = 10: Grd(2).Text = "Beta13": Grd(2).Row = 11: Grd(2).Text = "Beta23"
'make the scattering factors table
    chkIntTable_Click
    
    
'set the background grid, index=3

    Grd(3).Col = 0
    For i = 1 To Grd(3).Rows - 1
    Grd(3).Row = i
    Grd(3).CellAlignment = 0
    Grd(3).Text = CStr(i)
    Next i
    Grd(3).Col = 1: Grd(3).Row = 0: Grd(3).Text = "2 theta /deg"
    Grd(3).Col = 2: Grd(3).Row = 0: Grd(3).Text = "back. int."
    Grd(3).ColWidth(0) = 1.5 * Grd(3).Width / 6
    Grd(3).ColWidth(1) = 2 * Grd(3).Width / 6
    Grd(3).ColWidth(2) = 2 * Grd(3).Width / 6




RietveldBoardMessage strLinie
RietveldBoardMessage Now
RietveldBoardMessage "Warning: no checking for consistency of the input data will be made..."
RietveldBoardMessage "IMPORTANT: Any empty field will be taken either as zero ('0') for expected variables or as NULL for strings.  It is not necessary to fill with 0 the empty boxes."
RietveldBoardMessage strLinie
'open the config file
Call ControlFileOpen(App.Path & "\dbws\pw_ControlFile.cfg", coderoare)
If coderoare Then Err.Raise 1101

''hscPhase_Change
Exit Sub
errortrap:
RietveldBoardMessage "Error in opening the config file (pw_controlFile.cfg). " & Err.Description
Exit Sub
End Sub


Sub Form_Unload(Cancel As Integer)

If Not (CloseWindow("Are you sure you want to close this window ?", prog_name & " Control File")) Then Cancel = -1
End Sub

Private Sub Grd_Click(Index As Integer)
'grd(3) is the one for the user background
'2 is for the atom parameters
'1 is for scattering factors
'0 is for excluded regions
If Grd(Index).Row >= Grd(Index).Rows - 3 Then
Grd(Index).Rows = Grd(Index).Rows + 1
End If
End Sub

Private Sub Grd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo error_handler
'''If grd(index).Row = 0 Or grd(index).Col = 0 Then Exit Sub

Select Case KeyCode
Case 8 'backspace
If Len(Grd(Index).Text) > 0 Then Grd(Index).Text = left$(Grd(Index).Text, Len(Grd(Index).Text) - 1)
Case 13 'enter
''adjust here for each of the grids

If Grd(Index).Col = (Grd(Index).Cols - 1) Then
Grd(Index).Col = 1
If Grd(Index).Row > Grd(Index).Rows - 2 Then Grd(Index).Rows = Grd(Index).Rows + 1
Grd(Index).Row = Grd(Index).Row + 1
Else
Grd(Index).Col = Grd(Index).Col + 1
'''Call Grd
End If

Case 46 'delete
If Len(Grd(Index).Text) > 0 Then Grd(Index).Text = right$(Grd(Index).Text, Len(Grd(Index).Text) - 1)
Case 188, 110, 190 ', .
Grd(Index).Text = Grd(Index).Text + "."
Case 189, 109 '-
Grd(Index).Text = "-" & Grd(Index).Text
Case 48, 96 '0
Grd(Index).Text = Grd(Index).Text + "0"
Case 49, 97 '1
Grd(Index).Text = Grd(Index).Text + "1"
Case 50, 98 '2
Grd(Index).Text = Grd(Index).Text + "2"
Case 51, 99 '3
Grd(Index).Text = Grd(Index).Text + "3"
Case 52, 100 '4
Grd(Index).Text = Grd(Index).Text + "4"
Case 53, 101 '5
Grd(Index).Text = Grd(Index).Text + "5"
Case 54, 102 '6
Grd(Index).Text = Grd(Index).Text + "6"
Case 55, 103 '7
Grd(Index).Text = Grd(Index).Text + "7"
Case 56, 104 '8
Grd(Index).Text = Grd(Index).Text + "8"
Case 57, 105 '9
Grd(Index).Text = Grd(Index).Text + "9"

Case Else

'''MsgBox CStr(KeyCode)
End Select

Exit Sub
error_handler:
MsgBox "error in grid"
Err.Clear
Resume Next
End Sub

Private Sub hsc2_Change()
Call SetScatteringFactors
lblscatt.Caption = "set no: " & CInt(CStr(hsc2.Value))
CurrentScatteringSet = CInt(hsc2.Value)
Call FormScatteringFactors
End Sub

 Sub hscPhase_Change()
'phaseedited is the global parameter designating the last phase number edited
Me.MousePointer = 11
Me.Refresh
Call refreshPhaseData
DoEvents
Call refreshAtomData
DoEvents
lblPhaseSet.Caption = "PHASE " & CInt(CStr(hscPhase.Value))
phaseEdited = CInt(CStr(hscPhase.Value))
Call FormRefreshPhaseData
Call FormRefreshAtomData
Me.MousePointer = 0
Me.Refresh
DoEvents
'refresh all the data of former phase number phaseedit and last atom 'invoke the change of the scroll
'this refresh should be done also when exiting this form
End Sub


Private Sub mDataInMemory_Click()
On Error GoTo errorTR:


If numarvalori < 1 Then
MsgBox "There is no data loaded in memory.  You can proceed to the main window of the program and load a data file."
mDataInMemory.Checked = False

Else
mDataInMemory.Checked = True
mSelectOnDisk.Checked = False
'save as aDBWS file for DBWS
'check that the input file in set control  is as DBWS
Dim returncode As Boolean, i As Long, ultimalinie As String
On Error GoTo errortrap
raport strLinie
raport "Output in DBWS format."
Call verificadate(True, False, True, returncode)
If Not (returncode) Then MsgBox " Error: no data in memory ?": Exit Sub
outfil = FreeFile

'Call open_file(outputfile, 2, returncode)
'If Not (returncode) Then Exit Sub
'raport "The file is " & outputfile
outputfile = App.Path & "\dbws\_memDUMP.dat"

Open outputfile For Output As outfil
Print #outfil, Format$(Format$(startx, "#0.000##"), "@@@@@@@@"); Format$(Format$(stepx, "#0.000##"), "@@@@@@@@"); Format$(Format$(endx, "#0.000##"), "@@@@@@@@"), left$(title, 30)
For i = 1 To Fix(numarvalori / 8)
Print #outfil, Format$(Format$(Val(Y((i - 1) * 8 + 1)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 2)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 3)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 4)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 5)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 6)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 7)), "#####0. "), "@@@@@@@@") + Format$(Format$(Val(Y((i - 1) * 8 + 8)), "#####0. "), "@@@@@@@@")
Next i
'scriu eventualele puncte ramase...
If 8 * Fix(numarvalori / 8) < numarvalori Then
'au mai ramas cateva puncte...
ultimalinie = ""
For i = 8 * Fix(numarvalori / 8) + 1 To numarvalori
ultimalinie = ultimalinie + Format$(Format$(Val(Y(i)), "#####0. "), "@@@@@@@@")
Next i
Print #outfil, ultimalinie
End If
Close
'raport "DBWS file written, it seems to be ok...Check and adjust the output data to your needs."
'raport Now
Exit Sub
errortrap:
mDataInMemory.Checked = False
Err.Clear
'raport "An error has occured."
Close
Exit Sub
End If

Exit Sub
errorTR:
mDataInMemory.Checked = False
Exit Sub
End Sub

Private Sub mnuCFExit_Click()
Unload Me
End Sub

Private Sub mnuCFOpen_Click()
'here set up a control file with a generic name, let s say
Dim return_code As Boolean
Call open_file(inputfile, 1, return_code)
If Not (return_code) Then inputfile = "": Exit Sub
'dimension parts here
return_code = False
Call ControlFileOpen(inputfile, return_code)
If return_code Then MsgBox "error: " & Err.Description
Exit Sub
End Sub

Sub ControlFileOpen(numefile As String, eroare As Boolean)
On Error GoTo errortrap:

Dim i As Integer
Dim tTitle As String, tJobtyp As Integer, tNprof As Integer, tNphase As Integer, tNBCkgd As Integer
Dim tNexcrg As Integer, tNscat As Integer, tInstrm As Integer, tIpref As Integer, tIasym As Integer, tIAbsr As Integer, tIdata As Integer
Dim tIAS As Integer, tFondo As Integer, tTwentyFlags(20) As Integer, tWave1 As Single, twave2 As Single, tRatio As Single
Dim tBKpos As Single, tWDt As Single, tCTHM As Single, tTMR As Single, tRlim As Single, tSw As Single
Dim tMcycle As Integer, tEps As Single, tRelax(4) As Single, tThmin As Single, tThmax As Single, tthStep As Single
Dim tPOs() As Single, tBck() As Single, tAlow() As Single, tAhigh() As Single
Dim tMaxs As Integer 'number of parameters to be refined
Dim tZer As Single, tDisp As Single, tTrans As Single, tP As Single, tQ As Single, tR As Single, tt As Single
Dim tFlzer As Single, tFldisp As Single, tFltrns As Single, tFlgp As Single, tFlgq As Single, tFlgr As Single, tFlgt As Single
Dim tAm As Single, tMon1 As Single, tMon2 As Single, tFlam As Single, tFlmon1 As Single, tFlmon2 As Single, tBack(6) As Single, tFBack(6) As Single
Dim tPhsnm As String, tN As Integer, tFu As Integer, tAFQPA As Single, tPref(3) As Single
Dim tSYMB As String, tSf As Single, tBo As Single, tCs As Single, tCbo As Single
Dim sDummy(9) As Single, tNam As String * 4, tDFP As Single, tDFPP As Single, tAW As Single
Dim sLine As String, j As Integer

inpfil = FreeFile
Open numefile For Input As #inpfil
Line Input #inpfil, tTitle
Text4.Text = tTitle
'''Input #inpfil, tJobtyp, tNprof, tNphase, tNBCkgd, tNexcrg, tNscat, tInstrm, tIpref, tIasym, tibasr, tIdata
Line Input #inpfil, sLine '11i4
tJobtyp = CInt(Val(Mid$(sLine, 1, 4)))
tNprof = CInt(Val(Mid$(sLine, 5, 4)))
tNphase = CInt(Val(Mid$(sLine, 9, 4)))
tNBCkgd = CInt(Val(Mid$(sLine, 13, 4)))
tNexcrg = CInt(Val(Mid$(sLine, 17, 4)))
tNscat = CInt(Val(Mid$(sLine, 21, 4)))
tInstrm = CInt(Val(Mid$(sLine, 25, 4)))
tIpref = CInt(Val(Mid$(sLine, 29, 4)))
tIasym = CInt(Val(Mid$(sLine, 33, 4)))
tIAbsr = CInt(Val(Mid$(sLine, 37, 4)))
tIdata = CInt(Val(Mid$(sLine, 41, 4)))

Select Case tJobtyp
Case 0
chkJobType.Value = Unchecked
Check2(2).Value = Unchecked
Case 1
chkJobType.Value = Unchecked
Check2(2).Value = Checked
Case 2
chkJobType.Value = Checked
Check2(2).Value = Unchecked
Case 3
chkJobType.Value = Checked
Check2(2).Value = Checked
End Select
If tNBCkgd = -1 Then
Input #inpfil, tIAS, tFondo

OPTb(0).Value = True
Check1(0).Value = Unchecked
If tIAS = 1 Then Check1(0).Value = Checked
If tFondo = 1 Then OPTb(1).Value = True
If tFondo = 2 Then OPTb(2).Value = True
End If

lstProfileFunction.ListIndex = tNprof
txtNrPhases.Text = tNphase
    Select Case tNBCkgd
        Case 0
        optSetControlBackground(0).Value = True
        'normal
        Case 1
        optSetControlBackground(1).Value = True
        'file tape 3
        Case -1
        'riello
        optSetControlBackground(3).Value = True
        Case Else
        'linear interpolation
        optSetControlBackground(2).Value = True
        'see how many points, put them later in grd(3)
        Grd(3).Rows = tNBCkgd + 2
    End Select
txtExcl.Text = CStr(tNexcrg)
If tInstrm = 1 Then
chkVariableInt.Value = Checked
Else
chkVariableInt.Value = Unchecked
End If
If tIpref = 1 Then
Option1(1).Value = True
Else
Option1(1).Value = False
End If
If tIasym = 1 Then
Option2(1).Value = True
Else
Option2(1).Value = False
End If
Select Case tIAbsr
Case 1
Option3(0) = Checked
Case 2
Option3(1) = Checked
Case 3
Option3(2) = Checked
Case 4
Option3(3) = Checked
End Select
lstDataFormat.ListIndex = tIdata


Line Input #inpfil, sLine
For i = 1 To 5
tTwentyFlags(i) = CInt(Val(Mid$(sLine, i, 1)))
Next i
For i = 6 To 10
tTwentyFlags(i) = CInt(Val(Mid$(sLine, i + 1, 1)))
Next i

For i = 11 To 15
tTwentyFlags(i) = CInt(Val(Mid$(sLine, i + 2, 1)))
Next i

For i = 16 To 18
tTwentyFlags(i) = CInt(Val(Mid$(sLine, i + 3, 1)))
Next i

For i = 0 To 12: Check5(i).Value = Unchecked: Next i
For i = 0 To 6: Check6(i).Value = Unchecked: Next i
If tTwentyFlags(1) = 1 Then Check5(0).Value = Checked
If tTwentyFlags(2) = 1 Then Check5(1).Value = Checked
If tTwentyFlags(3) = 1 Then Check5(2).Value = Checked
If tTwentyFlags(3) = 2 Then Check5(3).Value = Checked
If tTwentyFlags(3) = 3 Then Check5(4).Value = Checked
If tTwentyFlags(4) = 1 Then Check5(5).Value = Checked
If tTwentyFlags(5) = 1 Then Check5(6).Value = Checked
If tTwentyFlags(6) = 1 Then Check5(7).Value = Checked
If tTwentyFlags(7) = 1 Then Check5(8).Value = Checked
If tTwentyFlags(8) = 1 Then Check5(9).Value = Checked
If tTwentyFlags(9) = 1 Then Check5(10).Value = Checked
tTwentyFlags(10) = 1 'not used, always 1
If tTwentyFlags(11) = 1 Then Check5(11).Value = Checked
If tTwentyFlags(11) = 2 Then Check5(12).Value = Checked
If tTwentyFlags(12) = 1 Then Check6(0).Value = Checked
If tTwentyFlags(13) = 1 Then Check6(1).Value = Checked
If tTwentyFlags(14) = 1 Then Check6(2).Value = Checked
If tTwentyFlags(15) = 1 Then Check6(3).Value = Checked
If tTwentyFlags(16) = 1 Then Check6(4).Value = Checked
If tTwentyFlags(17) = 1 Then Check6(5).Value = Checked
If tTwentyFlags(18) = 1 Then Check6(6).Value = Checked

'''Input #inpfil, tWave1, twave2, tRatio, tBKpos, tWDt, tCTHM, tTMR, tRlim, tSw
Line Input #inpfil, sLine
tWave1 = Val(Mid$(sLine, 1, 8))
twave2 = Val(Mid$(sLine, 9, 8))
tRatio = Val(Mid$(sLine, 17, 8))
tBKpos = Val(Mid$(sLine, 25, 8))
tWDt = Val(Mid$(sLine, 33, 8))
tCTHM = Val(Mid$(sLine, 41, 8))
tTMR = Val(Mid$(sLine, 49, 8))
tRlim = Val(Mid$(sLine, 57, 8))
tSw = Val(Mid$(sLine, 73, 8))


txtFixedPar(0).Text = CStr(Val(tWave1))
txtFixedPar(1).Text = CStr(Val(twave2))
txtFixedPar(2).Text = CStr(Val(tRatio))
txtFixedPar(3).Text = CStr(Val(tBKpos))
txtFixedPar(4).Text = CStr(Val(tWDt))
txtFixedPar(5).Text = CStr(Val(tCTHM))
txtFixedPar(6).Text = CStr(Val(tTMR))
txtFixedPar(7).Text = CStr(Val(tRlim))
txtFixedPar(8).Text = CStr(Val(tSw))
''Input #inpfil, tMcycle, tEps, tRelax(1), tRelax(2), tRelax(3), tRelax(4)

Line Input #inpfil, sLine
    tMcycle = Val(Mid$(sLine, 1, 4))
    tEps = Val(Mid$(sLine, 5, 4))
    tRelax(1) = Val(Mid$(sLine, 9, 4))
    tRelax(2) = Val(Mid$(sLine, 13, 4))
    tRelax(3) = Val(Mid$(sLine, 17, 4))
    tRelax(4) = Val(Mid$(sLine, 21, 4))
    txtOperations(0).Text = CStr(tMcycle)
    txtOperations(1).Text = CStr(tEps)
    txtOperations(2).Text = CStr(tRelax(1))
    txtOperations(3).Text = CStr(tRelax(2))
    txtOperations(4).Text = CStr(tRelax(3))
    txtOperations(5).Text = CStr(tRelax(4))

If tJobtyp = 2 Or tJobtyp = 3 Then
    '''Input #inpfil, tThmin, tstep, tThmax
    tThmin = Val(Mid$(sLine, 25, 8))
    tstep = Val(Mid$(sLine, 33, 8))
    tThmax = Val(Mid$(sLine, 41, 8))
    txtOperations(6).Text = CStr(tThmin)
    txtOperations(7).Text = CStr(tThmax)
    txtOperations(8).Text = CStr(tstep)
End If

If tNBCkgd > 2 Then
    ReDim tPOs(tNBCkgd), tBck(tNBCkgd)
    For i = 1 To tNBCkgd
    'read and put them in grid 3
        Grd(3).Rows = tNBCkgd + 3
        Line Input #inpfil, sLine
        Call SetGridVal(Grd(3), i, 1, Mid$(sLine, 1, 8))
        Call SetGridVal(Grd(3), i, 2, Mid$(sLine, 9, 8))
    Next i
End If

If tNexcrg > 0 Then
    ReDim tAlow(tNexcrg), tAhigh(tNexcrg)
    Grd(0).Rows = tNexcrg + 2
    For i = 1 To tNexcrg
        Line Input #inpfil, sLine
        Call SetGridVal(Grd(0), i, 0, Mid$(sLine, 1, 8))
        Call SetGridVal(Grd(0), i, 1, Mid$(sLine, 9, 8))
    Next i
End If

'read here the scattering factors, need nscat lines.......
ScattSetCallFromOpen = False
If tNscat > 0 Then
    totalScatt = tNscat
    ScattSetCallFromOpen = True
    chkScatt.Value = Checked
    lblscatt.Caption = "set no: 1"
'If chkScatt.Value = 1 Then
's = InputBox("How many scattering factors sets to add (max 20) ?", "scattering factors", 1)
    If tNscat > 0 And tNscat < 21 Then
    chkScatt.Caption = "set no: " & CStr(CInt(Val((tNscat))))
    hsc2.Value = 1
    Grd(1).Enabled = True
    lblscatt.Enabled = True
    hsc2.Enabled = True
    hsc2.Max = CInt(Val((tNscat)))
'the four texts
    For i = 0 To 3: txtScat(i).Enabled = True: Next i
'which way to write these factors
    If tJobtyp = 1 Or tJobtyp = 3 Then
        MsgBox ("I can read and write only XRD scattering factors. For neutrons you will need to use manual editing of the file. Ignoring neutron scattering data...")
    For i = 1 To tNscat
        Line Input #inpfil, sLine  'dummy
    Next i
    Else
'read here the XRD values
    For i = 1 To tNscat


        Line Input #inpfil, sLine
        Scattering(i).Name = Mid$(sLine, 1, 4)
        Scattering(i).RePart = Val(Mid$(sLine, 5, 8))
        Scattering(i).ImPart = Val(Mid$(sLine, 13, 8))
        Scattering(i).AtWeight = Val(Mid$(sLine, 21, 8))
'''t = InputBox("Set no:" & CStr(i) & ". Are the scattering factors in International Format, 9 constanst on a line ? (1=No, 0=Yes)  ", prog_name & " - scattering", "1")

Line Input #inpfil, sLine
If Len(sLine) < 17 Or Val(Mid$(sLine, 17)) = 0 Then
'pos scatt
Scattering(i).PosScatt(1, 1) = CSng(Val(Mid$(sLine, 1, 8)))
Scattering(i).PosScatt(2, 1) = CSng(Val(Mid$(sLine, 9)))
Scattering(i).IntTable = False
For j = 2 To 100 'change scatt factors to 100
Line Input #inpfil, sLine
If CSng(Val(Mid$(sLine, 1, 8))) < 0 Then Exit For
    Scattering(i).PosScatt(1, j) = CSng(Val(Mid$(sLine, 1, 8)))
    Scattering(i).PosScatt(2, j) = CSng(Val(Mid$(sLine, 9)))
Next j




Else
'international table

Scattering(i).IntTable = True
    For j = 1 To 9
        Scattering(i).NineCoeff(j) = CSng(Val(Mid$(sLine, 1 + (j - 1) * 8, 8)))
    Next j

End If
    Next i 'end of tnscat cycle
    End If
Else
    Err.Raise 1101, , "incorrect number of scattering factors,...up to 20 accepted. abort..."
End If
'now write in the text the first set
    

CurrentScatteringSet = 1
If Scattering(CurrentScatteringSet).IntTable Then
chkIntTable.Value = Checked
Else

chkIntTable.Value = Unchecked
End If
txtScat(0).Text = Scattering(CurrentScatteringSet).Name
txtScat(1).Text = CStr(Scattering(CurrentScatteringSet).RePart)
txtScat(2).Text = CStr(Scattering(CurrentScatteringSet).ImPart)
txtScat(3).Text = CStr(Scattering(CurrentScatteringSet).AtWeight)

If txtScat(1).Text = "0" Then txtScat(1).Text = ""
If txtScat(2).Text = "0" Then txtScat(2).Text = ""
If txtScat(3).Text = "0" Then txtScat(3).Text = ""

If Scattering(CurrentScatteringSet).IntTable Then
'put the 9 coefficients in the string nineCoeff
For i = 1 To 9: Call SetGridVal(Grd(1), i - 1, 1, CStr(Scattering(CurrentScatteringSet).NineCoeff(i))): Next i
Else
'read pos, scatt up to 50 sets of values
Grd(1).Rows = 110
For i = 1 To 100
Call SetGridVal(Grd(1), i - 1, 1, CStr(Scattering(CurrentScatteringSet).PosScatt(1, i)))
Call SetGridVal(Grd(1), i - 1, 2, CStr(Scattering(CurrentScatteringSet).PosScatt(2, i)))
Next i
End If
End If
ScattSetCallFromOpen = False


Line Input #inpfil, sLine
tMaxs = CInt(Val(Mid$(sLine, 1, 8)))
Line Input #inpfil, sLine
tZer = Val(Mid$(sLine, 1, 8))
tDisp = Val(Mid$(sLine, 9, 8))
tTrans = Val(Mid$(sLine, 17, 8))
tP = Val(Mid$(sLine, 25, 8))
tQ = Val(Mid$(sLine, 33, 8))
tR = Val(Mid$(sLine, 41, 8))
tt = Val(Mid$(sLine, 49, 8))

Line Input #inpfil, sLine
tFlzer = Val(Mid$(sLine, 1, 8))
tFldisp = Val(Mid$(sLine, 9, 8))
tflTrans = Val(Mid$(sLine, 17, 8))
tFlgp = Val(Mid$(sLine, 25, 8))
tFlgq = Val(Mid$(sLine, 33, 8))
tFlgr = Val(Mid$(sLine, 41, 8))
tFlgt = Val(Mid$(sLine, 49, 8))

txtParam.Text = CStr(tMaxs)

txtGlobalRefine(0).Text = CStr(Val(tZer))
txtGlobalRefine(1).Text = CStr(Val(tDisp))
txtGlobalRefine(2).Text = CStr(Val(tTrans))
txtGlobalRefine(3).Text = CStr(Val(tP))
txtGlobalRefine(4).Text = CStr(Val(tQ))
txtGlobalRefine(5).Text = CStr(Val(tR))
txtGlobalRefine(6).Text = CStr(Val(tt))

txtGlobalRefine(7).Text = CStr(Val(tFlzer))
txtGlobalRefine(8).Text = CStr(Val(tFldisp))
txtGlobalRefine(9).Text = CStr(Val(tflTrans))
txtGlobalRefine(10).Text = CStr(Val(tFlgp))
txtGlobalRefine(11).Text = CStr(Val(tFlgq))
txtGlobalRefine(12).Text = CStr(Val(tFlgr))
txtGlobalRefine(13).Text = CStr(Val(tFlgt))

If tNBCkgd = -1 Then
'amorphous data here, riello code
Line Input #inpfil, sLine
txtGlobalRefine(14) = Mid$(sLine, 1, 8)
txtGlobalRefine(15) = Mid$(sLine, 9, 8)
txtGlobalRefine(16) = Mid$(sLine, 17, 8)

Line Input #inpfil, sLine
txtGlobalRefine(17) = Mid$(sLine, 1, 8)
txtGlobalRefine(18) = Mid$(sLine, 9, 8)
txtGlobalRefine(19) = Mid$(sLine, 17, 8)
End If


If tNBCkgd = 0 Then
Line Input #inpfil, sLine
For i = 1 To 6
tBack(i) = Val(Mid$(sLine, 1 + (i - 1) * 9, 9))
Next i

Line Input #inpfil, sLine
For i = 1 To 6
tFBack(i) = Val(Mid$(sLine, 1 + (i - 1) * 9, 9))
Next i
End If

txtGlobalRefine(14).Text = CStr(Val(tAm))
txtGlobalRefine(15).Text = CStr(Val(tMon1))
txtGlobalRefine(16).Text = CStr(Val(tMon2))
txtGlobalRefine(17).Text = CStr(Val(tFlam))
txtGlobalRefine(18).Text = CStr(Val(tFlmon1))
txtGlobalRefine(19).Text = CStr(Val(tFlmon2))

For i = 20 To 25
txtGlobalRefine(i).Text = CStr(Val(tBack(i - 19)))
txtGlobalRefine(i + 6).Text = CStr(Val(tFBack(i - 19)))
Next i

'here are the phase related data
Dim tempS As String
txtNrPhases.Text = CStr(CInt(Val(tNphase)))
For j = 1 To tNphase
'''Input #inpfil, Phases(j).Name
Line Input #inpfil, Phases(j).Name
Line Input #inpfil, sLine
Phases(j).Atomi = CInt(Val(Mid$(sLine, 1, 4)))
Phases(j).FormulaUnits = CInt(Val(Mid$(sLine, 5, 4)))
Phases(j).ParticleAbsorptionFactor = Val(Mid$(sLine, 9, 8))
Phases(j).PrefOrientation(1) = Val(Mid$(sLine, 17, 4))
Phases(j).PrefOrientation(2) = Val(Mid$(sLine, 21, 4))
Phases(j).PrefOrientation(3) = Val(Mid$(sLine, 25, 4))

Line Input #inpfil, sLine
Phases(j).SpaceGroupSymbol = Mid$(sLine, 1, 20)
For i = 1 To Phases(j).Atomi

Line Input #inpfil, sLine
Atoms(i, j).Label = Mid$(sLine, 1, 4)
Atoms(i, j).Multiplicity = CInt(Val(Mid$(sLine, 6, 4)))
Atoms(i, j).Ntyp = ((Mid$(sLine, 11, 4)))
Atoms(i, j).X = (Val(Mid$(sLine, 17, 8)))
Atoms(i, j).Y = (Val(Mid$(sLine, 25, 8)))
Atoms(i, j).z = (Val(Mid$(sLine, 33, 8)))
Atoms(i, j).IsotropicThermal = (Val(Mid$(sLine, 41, 8)))
Atoms(i, j).SiteOccupancy = (Val(Mid$(sLine, 49, 8)))

Line Input #inpfil, sLine
Atoms(i, j).codeX = (Val(Mid$(sLine, 17, 8)))
Atoms(i, j).codeY = (Val(Mid$(sLine, 25, 8)))
Atoms(i, j).codeZ = (Val(Mid$(sLine, 33, 8)))
Atoms(i, j).codeIsotropicThermal = (Val(Mid$(sLine, 41, 8)))
Atoms(i, j).codeSiteOccupancy = (Val(Mid$(sLine, 49, 8)))

Line Input #inpfil, sLine
Atoms(i, j).Beta11 = (Val(Mid$(sLine, 1, 8)))
Atoms(i, j).Beta22 = (Val(Mid$(sLine, 9, 8)))
Atoms(i, j).Beta33 = (Val(Mid$(sLine, 17, 8)))
Atoms(i, j).Beta12 = (Val(Mid$(sLine, 25, 8)))
Atoms(i, j).Beta13 = (Val(Mid$(sLine, 33, 8)))
Atoms(i, j).Beta23 = (Val(Mid$(sLine, 41, 8)))

Line Input #inpfil, sLine
Atoms(i, j).codeBeta11 = (Val(Mid$(sLine, 1, 8)))
Atoms(i, j).codeBeta22 = (Val(Mid$(sLine, 9, 8)))
Atoms(i, j).codeBeta33 = (Val(Mid$(sLine, 17, 8)))
Atoms(i, j).codeBeta12 = (Val(Mid$(sLine, 25, 8)))
Atoms(i, j).codeBeta13 = (Val(Mid$(sLine, 33, 8)))
Atoms(i, j).codeBeta23 = (Val(Mid$(sLine, 41, 8)))
Next i

Line Input #inpfil, sLine
Phases(j).scalefactor = (Val(Mid$(sLine, 1, 8)))
Phases(j).OverallThermal = (Val(Mid$(sLine, 9, 8)))
Line Input #inpfil, sLine
Phases(j).scaleFactorCode = (Val(Mid$(sLine, 1, 8)))
Phases(j).OverallThermalCode = (Val(Mid$(sLine, 9, 8)))

Line Input #inpfil, sLine
Phases(j).U = (Val(Mid$(sLine, 1, 8)))
Phases(j).v = (Val(Mid$(sLine, 9, 8)))
Phases(j).W = (Val(Mid$(sLine, 17, 8)))
Phases(j).CT = (Val(Mid$(sLine, 25, 8)))
Phases(j).z = (Val(Mid$(sLine, 33, 8)))
Phases(j).X = (Val(Mid$(sLine, 41, 8)))
Phases(j).Y = (Val(Mid$(sLine, 49, 8)))

Line Input #inpfil, sLine
Phases(j).codeU = (Val(Mid$(sLine, 1, 8)))
Phases(j).codeV = (Val(Mid$(sLine, 9, 8)))
Phases(j).codeW = (Val(Mid$(sLine, 17, 8)))
Phases(j).codeCT = (Val(Mid$(sLine, 25, 8)))
Phases(j).codeZ = (Val(Mid$(sLine, 33, 8)))
Phases(j).codeX = (Val(Mid$(sLine, 41, 8)))
Phases(j).codeY = (Val(Mid$(sLine, 49, 8)))

Line Input #inpfil, sLine
Phases(j).a = (Val(Mid$(sLine, 1, 8)))
Phases(j).b = (Val(Mid$(sLine, 9, 8)))
Phases(j).c = (Val(Mid$(sLine, 17, 8)))
Phases(j).Alpha = (Val(Mid$(sLine, 25, 8)))
Phases(j).Beta = (Val(Mid$(sLine, 33, 8)))
Phases(j).gamma = (Val(Mid$(sLine, 41, 8)))

Line Input #inpfil, sLine
Phases(j).codeA = (Val(Mid$(sLine, 1, 8)))
Phases(j).codeB = (Val(Mid$(sLine, 9, 8)))
Phases(j).codeC = (Val(Mid$(sLine, 17, 8)))
Phases(j).codeAlpha = (Val(Mid$(sLine, 25, 8)))
Phases(j).codeBeta = (Val(Mid$(sLine, 33, 8)))
Phases(j).codegamma = (Val(Mid$(sLine, 41, 8)))

Line Input #inpfil, sLine
Phases(j).G1 = (Val(Mid$(sLine, 1, 8)))
Phases(j).G2 = (Val(Mid$(sLine, 9, 8)))
Phases(j).P = (Val(Mid$(sLine, 17, 8)))

Line Input #inpfil, sLine
Phases(j).codeG1 = (Val(Mid$(sLine, 1, 8)))
Phases(j).codeG2 = (Val(Mid$(sLine, 9, 8)))
Phases(j).codeP = (Val(Mid$(sLine, 17, 8)))

Line Input #inpfil, sLine
Phases(j).NA = (Val(Mid$(sLine, 1, 8)))
Phases(j).NB = (Val(Mid$(sLine, 9, 8)))
Phases(j).NC = (Val(Mid$(sLine, 17, 8)))

Line Input #inpfil, sLine
Phases(j).codeNA = (Val(Mid$(sLine, 1, 8)))
Phases(j).codeNB = (Val(Mid$(sLine, 9, 8)))
Phases(j).codeNC = (Val(Mid$(sLine, 17, 8)))

Line Input #inpfil, sLine
Phases(j).hNA = (Val(Mid$(sLine, 1, 8)))
Phases(j).hNB = (Val(Mid$(sLine, 9, 8)))
Phases(j).hNC = (Val(Mid$(sLine, 17, 8)))

Line Input #inpfil, sLine
Phases(j).codehNA = (Val(Mid$(sLine, 1, 8)))
Phases(j).codehNB = (Val(Mid$(sLine, 9, 8)))
Phases(j).codehNC = (Val(Mid$(sLine, 17, 8)))

Line Input #inpfil, sLine
Phases(j).SP7A = (Val(Mid$(sLine, 1, 8)))

Line Input #inpfil, sLine
Phases(j).codeSP7A = (Val(Mid$(sLine, 1, 8)))
Next j

Close #inpfil
Call FormRefreshAtomData
Call FormRefreshPhaseData
ScattSetCallFromOpen = False
tabControlFile_Click
Exit Sub
errortrap:
eroare = True
ScattSetCallFromOpen = False
MsgBox Err.Description

Close #inpfil

Exit Sub


End Sub


Private Sub mnuCFRun_Click()

'first save the control data file
On Error GoTo errortrap
'here set up a control file with a generic name
Call refreshAtomData
Call refreshPhaseData
Call SaveControlFile(App.Path & "\DBWS\_DB_runicf.dat")
Call SaveControlFile(App.Path & "\DBWS\backup_DB_runicf.dat")
'if the graph is open then close



Dim t As Double, sT As String
'ChDir App.Path & "\dbws"
dbwsDataFile = App.Path & "\dbws\_memDUMP.dat"
dbwsControlFile = App.Path & "\dbws\_DB_runICF.dat"
dbwsOutputFile = App.Path & "\dbws\_onLineRun.out"
If Not (left$(dbwsDataFile, 1) = """") Then dbwsDataFile = """" & dbwsDataFile & """"
If Not (left$(dbwsOutputFile, 1) = """") Then dbwsOutputFile = """" & dbwsOutputFile & """"
If Not (left$(dbwsControlFile, 1) = """") Then dbwsControlFile = """" & dbwsControlFile & """"

'sT = InputBox("This command will run the program pw_DBWS3.  You can change the filenames either manually or by selecting them with the Run/Select menu commands.  The order in this list is (i.e. datafile filename, control filename, output filename)", prog_name & " - shell DBWS98", dbwsDataFile & " " & dbwsControlFile & " " & dbwsOutputFile)
'If Len(sT) < 2 Then Exit Sub
sT = dbwsDataFile & " " & dbwsControlFile & " " & dbwsOutputFile

t = ShellAndLoop(App.Path & "\dbws\pw_DBWS3.exe " & sT, vbMaximizedFocus)
'RietveldBoardMessage "pw_DBWS3 called at " & Now
'RietveldBoardMessage strLinie





If mUpgradeCF.Checked Then

'here set up a control file with a generic name, let s say
Dim return_code As Boolean
'Call open_file(inputfile, 1, return_code)
'If Not (return_code) Then inputfile = "": Exit Sub
'dimension parts here
'return_code = False
Close
dbwsControlFile = App.Path & "\dbws\_DB_runICF.dat"
Call ControlFileOpen(dbwsControlFile, return_code)

End If




If mShowResults.Checked Then hWndShell "write.exe " & dbwsOutputFile, vbMinimizedNoFocus

'start the graph, to delete plotinfo ??
If mShowCFGraph.Checked Then


    inputfile = App.Path & "\dbws\plotinfo"
    
    newGraph.Show
    Call newGraph.fileopenplotinfo(inputfile)
End If

'upgrade the ICF file

Exit Sub

errortrap:
RietveldBoardMessage Err.Description
RietveldBoardMessage Now
RietveldBoardMessage strLinie
Err.Clear
Exit Sub






End Sub

Private Sub mnuEmpty_Click()
On Error GoTo errortrap
If (CloseWindow("Are you sure you want to clear the Control file ?", prog_name & " DBWS Control")) Then
'in this case closewindow can be used as well
'clear all the data from ControlFile containers














End If
Exit Sub

errortrap:
Exit Sub

End Sub

Private Sub mnuSaveControlFile_Click()
On Error GoTo errortrap
Dim returncode As Boolean
Call open_file(outputfile, 2, returncode)
If Not (returncode) Then Exit Sub
Call refreshAtomData
Call refreshPhaseData
Call SaveControlFile(outputfile)

Exit Sub

errortrap:
MsgBox Err.Description

Exit Sub
End Sub

Private Sub mnuSetControlFile_Click()
On Error GoTo errortrap
'here set up a control file with a generic name

Call refreshAtomData
Call refreshPhaseData
Call SaveControlFile(App.Path & "\DBWS\pw_icf.txt")
Exit Sub

errortrap:
MsgBox Err.Description
Close
Exit Sub
End Sub

Sub SaveControlFile(withFilename As String)
'withFilename should have already the path
't stands for temporary
Dim i As Integer
Dim tTitle As String, tJobtyp As Integer, tNprof As Integer, tNphase As Integer, tNBCkgd As Integer
Dim tNexcrg As Integer, tNscat As Integer, tInstrm As Integer, tIpref As Integer, tIasym As Integer, tIAbsr As Integer, tIdata As Integer
Dim tIAS As Integer, tFondo As Integer, tTwentyFlags(20) As Integer, tWave1 As Single, twave2 As Single, tRatio As Single
Dim tBKpos As Single, tWDt As Single, tCTHM As Single, tTMR As Single, tRlim As Single, tSw As Single
Dim tMcycle As Integer, tEps As Single, tRelax(4) As Single, tThmin As Single, tThmax As Single, tthStep As Single
Dim tPOs As Single, tBck As Single, tAlow As Single, tAhigh As Single
Dim tMaxs As Integer 'number of parameters to be refined
Dim tZer As Single, tDisp As Single, tTrans As Single, tP As Single, tQ As Single, tR As Single, tt As Single
Dim tFlzer As Single, tFldisp As Single, tFltrns As Single, tFlgp As Single, tFlgq As Single, tFlgr As Single, tFlgt As Single
Dim tAm As Single, tMon1 As Single, tMon2 As Single, tFlam As Single, tFlmon1 As Single, tFlmon2 As Single, tBack(6) As Single, tFBack(6) As Single
Dim tPhsnm As String, tN As Integer, tFu As Integer, tAFQPA As Single, tPref(3) As Single
Dim tSYMB As String, tSf As Single, tBo As Single, tCs As Single, tCbo As Single
'''On Error GoTo errortrap
Err.Clear
f = FreeFile
'update scatering factors
Call SetScatteringFactors
Call FormScatteringFactors
''Close #f

Open withFilename For Output As #f
tTitle = Mid$(Text4.Text, 1, 70) 'title
'next line is 10i4 ; 0,1,2,3: x, neutron, patt X, patt N
'''Print #f, tTitle
Print #f, sForFormat(tTitle, "A70")
tJobtyp = 2
If Check2(2).Value = Checked Then tJobtyp = 3
If chkJobType.Value = Unchecked Then tJobtyp = tJobtyp - 2
tNprof = lstProfileFunction.ListIndex
tNphase = CInt(Val(txtNrPhases.Text))
tNBCkgd = 0
'optSetControlBackground : 0=polynom, 1 file, 2=linear interpolation, 3=Riello
If optSetControlBackground(1).Value Then
MsgBox "Background should be prepared in a file Tape 3.  See the DBWS instruction manual"
tNBCkgd = 1
Else
    If optSetControlBackground(2).Value Then
    'linear int
    'see how many points ...in grd(3), stop when 2theta is 0
    For i = 1 To Grd(3).Rows - 2
    tNBCkgd = i - 1
    If Val(GridVal(Grd(3), i, 1)) <= 0 Then Exit For
    Next i
    If tNBCkgd < 2 Then
        MsgBox "Error: you need to input at least two sets of points for interpolation.  Background model control reset to 0": tNBCkgd = 0
    Else
        If optSetControlBackground(3).Value Then
            'Riello
            tNBCkgd = -1
        Else
            'nothing, stay zero
        End If
    End If
End If
End If
tNexcrg = CInt(Val(txtExcl.Text))
tNscat = 0
'to add here some code for the scattering
tNscat = totalScatt 'totalscatt must be smaller than 21

tInstrm = 0
If chkVariableInt.Value = Checked Then tInstrm = 1
'option1(0), 1 are for rietveld and March preffered orientation
tIpref = 0
If Option1(1).Value Then tIpref = 1 'march dollase function
tIasym = 0
If Option2(1).Value Then tIasym = 1 'Riello asymmetry
'option3 is for iabsr (surface roughness)
tIAbsr = 1
If Option3(1) Then
'sparks
tIAbsr = 2
Else
    If Option3(2) Then
    'suortti
    tIAbsr = 3
    Else
        If Option3(3) Then tIAbsr = 4
    End If
End If
'next is tIdata
tIdata = lstDataFormat.ListIndex ''
'''Print #f, Format$(Format$((tJobtyp), "###0"), "@@@@") & ", " & Format$(Format$((tNprof), "###0"), "@@@@") & ", " & Format$(Format$((tNphase), "###0"), "@@@@") & ", " & Format$(Format$((tNBCkgd), "###0"), "@@@@") & ", " & Format$(Format$((tNexcrg), "###0"), "@@@@") & ", " & Format$(Format$((tNscat), "###0"), "@@@@") & ", " & Format$(Format$((tInstrm), "###0"), "@@@@") & ", " & Format$(Format$((tIpref), "###0"), "@@@@") & ", " & Format$(Format$((tIasym), "###0"), "@@@@") & ", " & Format$(Format$((tIAbsr), "###0"), "@@@@") & ", " & Format$(Format$((tIdata), "###0"), "@@@@")
Print #f, sForFormat(tJobtyp, "i4") & sForFormat(tNprof, "i4") & sForFormat(tNphase, "i4") & sForFormat(tNBCkgd, "i4") & sForFormat(tNexcrg, "i4") & sForFormat(tNscat, "i4") & sForFormat(tInstrm, "i4") & sForFormat(tIpref, "i4") & sForFormat(tIasym, "i4") & sForFormat(tIAbsr, "i4") & sForFormat(tIdata, "i4")
If tNBCkgd = -1 Then
tIAS = 0: If Check1(0).Value = Checked Then tIAS = 1
'next is fondo
tFondo = 0
If OPTb(1).Value Then tFondo = 1
If OPTb(2).Value Then tFondo = 2
'''Print #f, Format$(Format$((tIAS), "###0"), "@@@@") & ", " & Format$(Format$((tFondo), "###0"), "@@@@")
Print #f, sForFormat(tIAS, "i4") & sForFormat(tFondo, "i4")
End If
    If Check5(0).Value = Checked Then tTwentyFlags(1) = 1
    If Check5(1).Value = Checked Then tTwentyFlags(2) = 1
    If Check5(2).Value = Checked Then tTwentyFlags(3) = 1
    If Check5(3).Value = Checked Then tTwentyFlags(3) = 2
    If Check5(4).Value = Checked Then tTwentyFlags(3) = 3
    If Check5(5).Value = Checked Then tTwentyFlags(4) = 1
    If Check5(6).Value = Checked Then tTwentyFlags(5) = 1
    If Check5(7).Value = Checked Then tTwentyFlags(6) = 1
    If Check5(8).Value = Checked Then tTwentyFlags(7) = 1
    If Check5(9).Value = Checked Then tTwentyFlags(8) = 1
    If Check5(10).Value = Checked Then tTwentyFlags(9) = 1
    tTwentyFlags(10) = 1 'not used, always 1
    If Check5(11).Value = Checked Then tTwentyFlags(11) = 1
    If Check5(12).Value = Checked Then tTwentyFlags(11) = 2
    If Check6(0).Value = Checked Then tTwentyFlags(12) = 1
    If Check6(1).Value = Checked Then tTwentyFlags(13) = 1
    If Check6(2).Value = Checked Then tTwentyFlags(14) = 1
    If Check6(3).Value = Checked Then tTwentyFlags(15) = 1
    If Check6(4).Value = Checked Then tTwentyFlags(16) = 1
    If Check6(5).Value = Checked Then tTwentyFlags(17) = 1
    If Check6(6).Value = Checked Then tTwentyFlags(18) = 1
Print #f, sForFormat(tTwentyFlags(1), "i1") & sForFormat(tTwentyFlags(2), "i1") & sForFormat(tTwentyFlags(3), "i1") & sForFormat(tTwentyFlags(4), "i1") & sForFormat(tTwentyFlags(5), "i1") & " " & sForFormat(tTwentyFlags(6), "i1") & sForFormat(tTwentyFlags(7), "i1") & sForFormat(tTwentyFlags(8), "i1") & sForFormat(tTwentyFlags(9), "i1") & sForFormat(tTwentyFlags(10), "i1") & " " & sForFormat(tTwentyFlags(11), "i1") & sForFormat(tTwentyFlags(12), "i1") & sForFormat(tTwentyFlags(13), "i1") & sForFormat(tTwentyFlags(14), "i1") & sForFormat(tTwentyFlags(15), "i1") & " " & sForFormat(tTwentyFlags(16), "i1") & sForFormat(tTwentyFlags(17), "i1") & sForFormat(tTwentyFlags(18), "i1")
tWave1 = CSng(Val(txtFixedPar(0).Text))
twave2 = CSng(Val(txtFixedPar(1).Text))
tRatio = CSng(Val(txtFixedPar(2).Text))
tBKpos = CSng(Val(txtFixedPar(3).Text))
tWDt = CSng(Val(txtFixedPar(4).Text))
tCTHM = CSng(Val(txtFixedPar(5).Text))
tTMR = CSng(Val(txtFixedPar(6).Text))
tRlim = CSng(Val(txtFixedPar(7).Text))
tSw = CSng(Val(txtFixedPar(8).Text))
Print #f, sForFormat(tWave1, "F8.0") & sForFormat(twave2, "F8.0") & sForFormat(tRatio, "F8.0") & sForFormat(tBKpos, "F8.0") & sForFormat(tWDt, "F8.0") & sForFormat(tCTHM, "F8.0") & sForFormat(tTMR, "F8.0") & sForFormat(tRlim, "F8.0") & sForFormat(tSw, "F8.0")
    tMcycle = CInt(Val(txtOperations(0).Text))
    tEps = CSng(Val(txtOperations(1).Text))
    tRelax(1) = CSng(Val(txtOperations(2).Text))
    tRelax(2) = CSng(Val(txtOperations(3).Text))
    tRelax(3) = CSng(Val(txtOperations(4).Text))
    tRelax(4) = CSng(Val(txtOperations(5).Text))
    tThmin = CSng(Val(txtOperations(6).Text))
    tThmax = CSng(Val(txtOperations(7).Text))
    tthStep = CSng(Val(txtOperations(8).Text))
Select Case tJobtyp
    Case 0, 1
        Print #f, sForFormat(tMcycle, "I4") & sForFormat(tEps, "F4.0") & sForFormat(tRelax(1), "F4.0") & sForFormat(tRelax(2), "F4.0") & sForFormat(tRelax(3), "F4.0") & sForFormat(tRelax(4), "F4.0")
    Case Else
        Print #f, sForFormat(tMcycle, "I4") & sForFormat(tEps, "F4.0") & sForFormat(tRelax(1), "F4.0") & sForFormat(tRelax(2), "F4.0") & sForFormat(tRelax(3), "F4.0") & sForFormat(tRelax(4), "F4.0") & sForFormat(tThmin, "F8.0") & sForFormat(tthStep, "F8.0") & sForFormat(tThmax, "F8.0")
    End Select
If tNBCkgd > 2 Then
'the four grids are: 3 for background, 2 for the atoms, 1 for scattering, 0 for the excluded regions
    For i = 1 To tNBCkgd
        tPOs = CSng(GridVal(Grd(3), i, 1))
        tBck = CSng(GridVal(Grd(3), i, 2))
    Print #f, sForFormat(tPOs, "F8.2") & sForFormat(tBck, "F8.2")
    Next i
End If
If tNexcrg > 0 Then
    For i = 1 To tNexcrg
        tAlow = CSng(Val(GridVal(Grd(0), i, 0)))
        tAhigh = CSng(Val(GridVal(Grd(0), i, 1)))
        If tAhigh = 0 Then MsgBox "Error in excluded region definition.  Possible cause: the number of excluded regions exceeds the number of data sets inserted in grid.  Check the data file.": Exit For
        Print #f, sForFormat(tAlow, "F8.2") & sForFormat(tAhigh, "F8.2")
        Next i
End If
If tNscat > 0 Then
'write here the scattering factors, need nscat lines
If chkScatt.Value = Checked Then
'add here the values for up to 20 scatt factors
'Type ScatteringFactors
'    IntTable As Boolean
'    Name As String
'    RePart As Single
'    ImPart As Single
'    AtWeight As Single
'    NineCoeff(9) As Single
'    PosScatt(2, 100) As Single
'End Type
'Global Scattering(20) As ScatteringFactors
If jobtyp = 0 Or jobtyp = 2 Then
'xrays
For i = 1 To tNscat
    Print #f, sForFormat(UCase$(Scattering(i).Name), "A4") & sForFormat(Scattering(i).RePart, "F8.0") & sForFormat(Scattering(i).ImPart, "F8.0") & sForFormat(Scattering(i).AtWeight, "F8.0")
    If Scattering(i).IntTable Then
            Print #f, sForFormat(Scattering(i).NineCoeff(1), "F8.0") & sForFormat(Scattering(i).NineCoeff(2), "F8.0") & sForFormat(Scattering(i).NineCoeff(3), "F8.0") & sForFormat(Scattering(i).NineCoeff(4), "F8.0") & sForFormat(Scattering(i).NineCoeff(5), "F8.0") & sForFormat(Scattering(i).NineCoeff(6), "F8.0") & sForFormat(Scattering(i).NineCoeff(7), "F8.0") & sForFormat(Scattering(i).NineCoeff(8), "F8.0") & sForFormat(Scattering(i).NineCoeff(9), "F8.0")
        Else
'    MsgBox "This program will calculate the 9 coefficients based on the data you entered (see the help file). Please check the output data file."
            
            
            For j = 1 To 100
If j > 30 Then MsgBox "DBWS accepts only 30 scattering sets.": Exit For
            '''MsgBox CStr(Scattering(i).PosScatt(1, j)) & " j= " & CStr(j)
            If (Scattering(i).PosScatt(1, j) <= 0 And (j > 1)) Then
                Print #f, "-100."
                Exit For
            Else
'attention: in the DBWS manual there is no format definition  for this case, check the source code
Print #f, sForFormat(Val(Scattering(i).PosScatt(1, j)), "F8.0") & sForFormat(Val(Scattering(i).PosScatt(2, j)), "F8.0") & "                                                                           "
            End If
            Next j
        
        End If
    Next i
Else
    'neutrons
    MsgBox "Neutron experiment: please insert the scattering length in the file"
    For i = 1 To tNscat
        Print #f, sForFormat(Scattering(i).Name, "A4")
        Print #f, " insert here the scattering length in units of 10**(-12)cm"
        Print #f, sForFormat(Scattering(i).AtWeight, "F8.0")
    Next i
End If
End If
End If
tMaxs = CInt(Val(txtParam.Text))
If tMaxs = 0 Then MsgBox "No refinement of the parameters ??? "
Print #f, sForFormat(tMaxs, "I8")
    tZer = CSng(Val(txtGlobalRefine(0).Text))
    tDisp = CSng(Val(txtGlobalRefine(1).Text))
    tTrans = CSng(Val(txtGlobalRefine(2).Text))
    tP = CSng(Val(txtGlobalRefine(3).Text))
    tQ = CSng(Val(txtGlobalRefine(4).Text))
    tR = CSng(Val(txtGlobalRefine(5).Text))
    tt = CSng(Val(txtGlobalRefine(6).Text))

    tFlzer = CSng(Val(txtGlobalRefine(7).Text))
    tFldisp = CSng(Val(txtGlobalRefine(8).Text))
    tflTrans = CSng(Val(txtGlobalRefine(9).Text))
    tFlgp = CSng(Val(txtGlobalRefine(10).Text))
    tFlgq = CSng(Val(txtGlobalRefine(11).Text))
    tFlgr = CSng(Val(txtGlobalRefine(12).Text))
    tFlgt = CSng(Val(txtGlobalRefine(13).Text))

    tAm = CSng(Val(txtGlobalRefine(14).Text))
    tMon1 = CSng(Val(txtGlobalRefine(15).Text))
    tMon2 = CSng(Val(txtGlobalRefine(16).Text))
    tFlam = CSng(Val(txtGlobalRefine(17).Text))
    tFlmon1 = CSng(Val(txtGlobalRefine(18).Text))
    tFlmon2 = CSng(Val(txtGlobalRefine(19).Text))
For i = 20 To 25
    tBack(i - 19) = CSng(Val(txtGlobalRefine(i).Text))
    tFBack(i - 19) = CSng(Val(txtGlobalRefine(i + 6).Text))
Next i
'global parameters
'6 lines
    Print #f, sForFormat(tZer, "F8.0") & sForFormat(tDisp, "F8.0") & sForFormat(tTrans, "F8.0") & sForFormat(tP, "F8.0") & sForFormat(tQ, "F8.0") & sForFormat(tR, "F8.0") & sForFormat(tt, "F8.0")
    Print #f, sForFormat(tFlzer, "F8.0") & sForFormat(tFldisp, "F8.0") & sForFormat(tflTrans, "F8.0") & sForFormat(tFlgp, "F8.0") & sForFormat(tFlgq, "F8.0") & sForFormat(tFlgr, "F8.0") & sForFormat(tFlgt, "F8.0")
If tNBCkgd = -1 Then
    Print #f, sForFormat(tAm, "F8.0") & sForFormat(tMon1, "F8.0") & sForFormat(tMon2, "F8.0")
    Print #f, sForFormat(tFlam, "F8.0") & sForFormat(tFlmon1, "F8.0") & sForFormat(tFlmon2, "F8.0")
End If
If tNBCkgd = 0 Then
    Print #f, sForFormat(tBack(1), "F9.4") & sForFormat(tBack(2), "F9.4") & sForFormat(tBack(3), "F9.4") & sForFormat(tBack(4), "F9.4") & sForFormat(tBack(5), "F9.4") & sForFormat(tBack(6), "F9.4")
    Print #f, sForFormat(tFBack(1), "F9.4") & sForFormat(tFBack(2), "F9.4") & sForFormat(tFBack(3), "F9.4") & sForFormat(tFBack(4), "F9.4") & sForFormat(tFBack(5), "F9.4") & sForFormat(tFBack(6), "F9.4")
End If
'>>>phase related data
For j = 1 To CInt(Val(txtNrPhases.Text))
    tPhsnm = CStr(Phases(j).Name)
    tN = CInt(Phases(j).Atomi)
    tFu = CInt(Phases(j).FormulaUnits)
    tAFQPA = CSng(Phases(j).ParticleAbsorptionFactor)
    tPref(1) = CSng(Phases(j).PrefOrientation(1))
    tPref(2) = CSng(Phases(j).PrefOrientation(2))
    tPref(3) = CSng(Phases(j).PrefOrientation(3))
tSYMB = CStr(Phases(j).SpaceGroupSymbol)
Print #f, sForFormat(tPhsnm, "A70")
Print #f, sForFormat(tN, "I4") & sForFormat(tFu, "I4") & sForFormat(tAFQPA, "F8.0") & sForFormat(tPref(1), "F4.0") & sForFormat(tPref(2), "F4.0") & sForFormat(tPref(3), "F4.0")
Print #f, sForFormat(UCase$(tSYMB), "A20")
For i = 1 To tN
    Print #f, sForFormat(Atoms(i, j).Label, "A4") & " " & sForFormat(Atoms(i, j).Multiplicity, "i4") & " " & sForFormat(Atoms(i, j).Ntyp, "A4") & "  " & sForFormat(Atoms(i, j).X, "F8.0") & sForFormat(Atoms(i, j).Y, "F8.0") & sForFormat(Atoms(i, j).z, "F8.0") & sForFormat(Atoms(i, j).IsotropicThermal, "F8.0") & sForFormat(Atoms(i, j).SiteOccupancy, "F8.0")
    Print #f, "                " & sForFormat(Atoms(i, j).codeX, "F8.0") & sForFormat(Atoms(i, j).codeY, "F8.0") & sForFormat(Atoms(i, j).codeZ, "F8.0") & sForFormat(Atoms(i, j).codeIsotropicThermal, "F8.0") & sForFormat(Atoms(i, j).codeSiteOccupancy, "F8.0")
    Print #f, sForFormat(Atoms(i, j).Beta11, "F8.0") & sForFormat(Atoms(i, j).Beta22, "F8.0") & sForFormat(Atoms(i, j).Beta33, "F8.0") & sForFormat(Atoms(i, j).Beta12, "F8.0") & sForFormat(Atoms(i, j).Beta13, "F8.0") & sForFormat(Atoms(i, j).Beta23, "F8.0")
    Print #f, sForFormat(Atoms(i, j).codeBeta11, "F8.0") & sForFormat(Atoms(i, j).codeBeta22, "F8.0") & sForFormat(Atoms(i, j).codeBeta33, "F8.0") & sForFormat(Atoms(i, j).codeBeta12, "F8.0") & sForFormat(Atoms(i, j).codeBeta13, "F8.0") & sForFormat(Atoms(i, j).codeBeta23, "F8.0")
Next i
    Print #f, sForFormat(Phases(j).scalefactor, "F8.0") & sForFormat(Phases(j).OverallThermal, "F8.0")
    Print #f, sForFormat(Phases(j).scaleFactorCode, "F8.0") & sForFormat(Phases(j).OverallThermalCode, "F8.0")
    Print #f, sForFormat(Phases(j).U, "F8.0") & sForFormat(Phases(j).v, "F8.0") & sForFormat(Phases(j).W, "F8.0") & sForFormat(Phases(j).CT, "F8.0") & sForFormat(Phases(j).z, "F8.0") & sForFormat(Phases(j).X, "F8.0") & sForFormat(Phases(j).Y, "F8.0")
    Print #f, sForFormat(Phases(j).codeU, "F8.0") & sForFormat(Phases(j).codeV, "F8.0") & sForFormat(Phases(j).codeW, "F8.0") & sForFormat(Phases(j).codeCT, "F8.0") & sForFormat(Phases(j).codeZ, "F8.0") & sForFormat(Phases(j).codeX, "F8.0") & sForFormat(Phases(j).codeY, "F8.0")
    Print #f, sForFormat(Phases(j).a, "F8.0") & sForFormat(Phases(j).b, "F8.0") & sForFormat(Phases(j).c, "F8.0") & sForFormat(Phases(j).Alpha, "F8.0") & sForFormat(Phases(j).Beta, "F8.0") & sForFormat(Phases(j).gamma, "F8.0")
    Print #f, sForFormat(Phases(j).codeA, "F8.0") & sForFormat(Phases(j).codeB, "F8.0") & sForFormat(Phases(j).codeC, "F8.0") & sForFormat(Phases(j).codeAlpha, "F8.0") & sForFormat(Phases(j).codeBeta, "F8.0") & sForFormat(Phases(j).codegamma, "F8.0")
    Print #f, sForFormat(Phases(j).G1, "F8.0") & sForFormat(Phases(j).G2, "F8.0") & sForFormat(Phases(j).P, "F8.0")
    Print #f, sForFormat(Phases(j).codeG1, "F8.0") & sForFormat(Phases(j).codeG2, "F8.0") & sForFormat(Phases(j).codeP, "F8.0")
    Print #f, sForFormat(Phases(j).NA, "F8.4") & sForFormat(Phases(j).NB, "F8.4") & sForFormat(Phases(j).NC, "F8.4")
    Print #f, sForFormat(Phases(j).codeNA, "F8.2") & sForFormat(Phases(j).codeNB, "F8.2") & sForFormat(Phases(j).codeNC, "F8.2")
    Print #f, sForFormat(Phases(j).hNA, "F8.4") & sForFormat(Phases(j).hNB, "F8.4") & sForFormat(Phases(j).hNC, "F8.4")
    Print #f, sForFormat(Phases(j).codehNA, "F8.2") & sForFormat(Phases(j).codehNB, "F8.2") & sForFormat(Phases(j).codehNC, "F8.2")
    Print #f, sForFormat(Phases(j).SP7A, "F8.4")
    Print #f, sForFormat(Phases(j).codeSP7A, "F8.2")
Next j
Close #f

RietveldBoardMessage " gDBWS file saved as " & withFilename & " " & Now
Beep
Exit Sub
errortrap:
Close #f
MsgBox "Error encountered in Sub SaveControlfile. " & Err.Description
RietveldBoardMessage strLinie
RietveldBoardMessage "Error encountered in Sub SaveControlfile. " & Err.Description & " error number " & Err.Number
RietveldBoardMessage "check the output file..."
Exit Sub
End Sub


Private Sub mSelectOnDisk_Click()
mSelectOnDisk.Checked = False

On Error GoTo error_open
Me.MousePointer = 11
'Select Case intrare_iesire
'    Case 1 'citirea path_input_text
Convert3Main.Dialog.Filter = "text file (*.txt) |*.txt|data file (*.dat) |*.dat|show all (*.*) |*.*"
Convert3Main.Dialog.FilterIndex = 3
Convert3Main.Dialog.Flags = &H1000& Or &H4& Or &H800&
'ofn_filemustexist 'ofn_readonly 'ofn_pathmustexist
Convert3Main.Dialog.DialogTitle = prog_name & " - select DBWS data file"
Convert3Main.Dialog.Action = 1
If Convert3Main.Dialog.FileName = App.Path & "\dbws\_memDUMP.dat" Then Err.Raise 1101
 Call CopyFile(Convert3Main.Dialog.FileName, App.Path & "\dbws\_memDUMP.dat")
'filesystemobject.Copyfile Convert3Main.Dialog.FileName, App.Path & "\dbws\_memDUMP.dat"
''Convert3Main.Dialog.FileName.Copy App.Path & "\dbws\_memDUMP.dat"

mSelectOnDisk.Checked = True
mDataInMemory.Checked = False
'dbwsDataFile = Convert3Main.Dialog.FileName

Me.MousePointer = 0

Exit Sub
error_open:
Err.Clear
dbwsDataFile = ""
mSelectOnDisk.Checked = False
Me.MousePointer = 0
Exit Sub

End Sub

Private Sub mShowCFGraph_Click()
mShowCFGraph.Checked = Not (mShowCFGraph.Checked)
End Sub

Private Sub mShowResults_Click()
mShowResults.Checked = Not (mShowResults.Checked)
End Sub

Private Sub mUpgradeCF_Click()
mUpgradeCF.Checked = Not (mUpgradeCF.Checked)
End Sub

Private Sub optSetControlBackground_Click(Index As Integer)
Dim i As Integer, return_code As Boolean
'delete all unnecessary parts
'txtBackgroundFile.Enabled = False
Grd(3).Enabled = False
Check1(0).Enabled = False
Frame7.Enabled = False
OPTb(0).Enabled = False: OPTb(1).Enabled = False: OPTb(2).Enabled = False

'now enable the necessary parts
Select Case Index
Case 1

Case 2

Grd(3).Enabled = True
Case 3
Check1(0).Enabled = True
Frame7.Enabled = True
OPTb(0).Enabled = True: OPTb(1).Enabled = True: OPTb(2).Enabled = True
End Select





End Sub





 Sub tabControlFile_Click()
Dim i As Integer
'erase all tabs
'in tab, the property .selected item give the caption of the selected tab
For i = 0 To 7 ''
frmControlFile.SetControlFrameTab(i).Visible = False
Next i
DoEvents
'show only the controls on the tab which is active
'(frmControlFile.tabControlFile.SelectedItem.Index) is the number of the tab shown
If frmControlFile.tabControlFile.SelectedItem.Index <= 7 Then
frmControlFile.SetControlFrameTab(CInt(-1 + (frmControlFile.tabControlFile.SelectedItem.Index))).Visible = True
Else
frmControlFile.SetControlFrameTab(7).Visible = True
End If
'''see how many tabs we need
'read the number of phases from txtnrphases
i = CInt(txtNrPhases.Text)
If i < 1 Then
txtNrPhases.Text = 1
Else
If i > 15 Then
txtNrPhases.Text = 15
End If
End If
'the number of phases
hscPhase.Max = Val(txtNrPhases.Text)
'for pattern calculation, the frame
Select Case chkJobType.Value
Case 0 'unchecked
frameRunChoices.Enabled = False
Case 1 'checked
frameRunChoices.Enabled = True
End Select

''here set at least one of the scatering factors (first and the last)
''the others are read and set by pressing the set no ScrollBar
Call SetScatteringFactors
'add on 13th of december
Call refreshAtomData
Call refreshPhaseData

End Sub

Private Sub txtExcl_Change()
If Val(txtExcl.Text) > 0 Then
Grd(0).Rows = 1 + Int(Val(txtExcl.Text))

End If
End Sub


Private Sub txtPhase_Change(Index As Integer)
'index=1 is for the number of atoms
If (Val(CStr(txtPhase(1).Text)) < 1) Then txtPhase(1).Text = 1
If (Val(CStr(txtPhase(1).Text)) > 150) Then txtPhase(1).Text = 200
Grd(2).Rows = Int(Val(CStr(txtPhase(1).Text)) * 25)
If Index = 1 Then
atScroll.Max = Val(CStr(txtPhase(1).Text))
'Frame6.Caption = "atom no: " & CInt(CStr(atScroll.Min))
End If

End Sub
