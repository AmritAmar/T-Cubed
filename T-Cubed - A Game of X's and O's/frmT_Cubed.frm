VERSION 5.00
Begin VB.Form frmT_Cubed 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T-Cubed - A Game of X's and O's"
   ClientHeight    =   8820
   ClientLeft      =   7320
   ClientTop       =   2088
   ClientWidth     =   7572
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   7572
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Restart Game"
      Height          =   492
      Left            =   4080
      TabIndex        =   105
      Top             =   7920
      Width           =   1452
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   1440
      Top             =   7680
   End
   Begin VB.Frame TCubedBoard 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7452
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7332
      Begin VB.Frame MainBoard 
         BackColor       =   &H00E0E0E0&
         Height          =   2412
         Index           =   3
         Left            =   4920
         TabIndex        =   27
         Top             =   120
         Width           =   2292
         Begin VB.Label TileB3 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   9
            Left            =   1560
            TabIndex        =   36
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB3 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   8
            Left            =   840
            TabIndex        =   35
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB3 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   7
            Left            =   120
            TabIndex        =   34
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB3 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   6
            Left            =   1560
            TabIndex        =   33
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB3 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   5
            Left            =   840
            TabIndex        =   32
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB3 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   4
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB3 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   3
            Left            =   1560
            TabIndex        =   30
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB3 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   2
            Left            =   840
            TabIndex        =   29
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB3 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   612
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   612
         End
      End
      Begin VB.Frame MainBoard 
         BackColor       =   &H00E0E0E0&
         Height          =   2412
         Index           =   9
         Left            =   4920
         TabIndex        =   26
         Top             =   4920
         Width           =   2292
         Begin VB.Label TileB9 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   9
            Left            =   1560
            TabIndex        =   90
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB9 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   8
            Left            =   840
            TabIndex        =   89
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB9 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   7
            Left            =   120
            TabIndex        =   88
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB9 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   6
            Left            =   1560
            TabIndex        =   87
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB9 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   5
            Left            =   840
            TabIndex        =   86
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB9 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   4
            Left            =   120
            TabIndex        =   85
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB9 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   3
            Left            =   1560
            TabIndex        =   84
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB9 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   2
            Left            =   840
            TabIndex        =   83
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB9 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   612
            Index           =   1
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   612
         End
      End
      Begin VB.Frame MainBoard 
         BackColor       =   &H00E0E0E0&
         Height          =   2412
         Index           =   8
         Left            =   2520
         TabIndex        =   25
         Top             =   4920
         Width           =   2292
         Begin VB.Label TileB8 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   9
            Left            =   1560
            TabIndex        =   81
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB8 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   8
            Left            =   840
            TabIndex        =   80
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB8 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   7
            Left            =   120
            TabIndex        =   79
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB8 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   6
            Left            =   1560
            TabIndex        =   78
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB8 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   5
            Left            =   840
            TabIndex        =   77
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB8 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   4
            Left            =   120
            TabIndex        =   76
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB8 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   3
            Left            =   1560
            TabIndex        =   75
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB8 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   2
            Left            =   840
            TabIndex        =   74
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB8 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   1
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   612
         End
      End
      Begin VB.Frame MainBoard 
         BackColor       =   &H00E0E0E0&
         Height          =   2412
         Index           =   7
         Left            =   120
         TabIndex        =   24
         Top             =   4920
         Width           =   2292
         Begin VB.Label TileB7 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   9
            Left            =   1560
            TabIndex        =   72
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB7 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   8
            Left            =   840
            TabIndex        =   71
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB7 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   7
            Left            =   120
            TabIndex        =   70
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB7 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   6
            Left            =   1560
            TabIndex        =   69
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB7 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   5
            Left            =   840
            TabIndex        =   68
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB7 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   4
            Left            =   120
            TabIndex        =   67
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB7 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   3
            Left            =   1560
            TabIndex        =   66
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB7 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   2
            Left            =   840
            TabIndex        =   65
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB7 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   612
            Index           =   1
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   612
         End
      End
      Begin VB.Frame MainBoard 
         BackColor       =   &H00E0E0E0&
         Height          =   2412
         Index           =   6
         Left            =   4920
         TabIndex        =   23
         Top             =   2520
         Width           =   2292
         Begin VB.Label TileB6 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   9
            Left            =   1560
            TabIndex        =   63
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB6 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   8
            Left            =   840
            TabIndex        =   62
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB6 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   7
            Left            =   120
            TabIndex        =   61
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB6 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   6
            Left            =   1560
            TabIndex        =   60
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB6 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   5
            Left            =   840
            TabIndex        =   59
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB6 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   4
            Left            =   120
            TabIndex        =   58
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB6 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   3
            Left            =   1560
            TabIndex        =   57
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB6 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   2
            Left            =   840
            TabIndex        =   56
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB6 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   612
            Index           =   1
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   612
         End
      End
      Begin VB.Frame MainBoard 
         BackColor       =   &H00E0E0E0&
         Height          =   2412
         Index           =   5
         Left            =   2520
         MousePointer    =   2  'Cross
         TabIndex        =   22
         Top             =   2520
         Width           =   2292
         Begin VB.Label TileB5 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   9
            Left            =   1560
            TabIndex        =   54
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB5 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   8
            Left            =   840
            TabIndex        =   53
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB5 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   7
            Left            =   120
            TabIndex        =   52
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB5 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   6
            Left            =   1560
            TabIndex        =   51
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB5 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   5
            Left            =   840
            TabIndex        =   50
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB5 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   4
            Left            =   120
            TabIndex        =   49
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB5 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   3
            Left            =   1560
            TabIndex        =   48
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB5 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   2
            Left            =   840
            TabIndex        =   47
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB5 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   612
            Index           =   1
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   612
         End
      End
      Begin VB.Frame MainBoard 
         BackColor       =   &H00E0E0E0&
         Height          =   2412
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   2292
         Begin VB.Label TileB4 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   9
            Left            =   1560
            TabIndex        =   45
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB4 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   8
            Left            =   840
            TabIndex        =   44
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB4 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   7
            Left            =   120
            TabIndex        =   43
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB4 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   6
            Left            =   1560
            TabIndex        =   42
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB4 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   5
            Left            =   840
            TabIndex        =   41
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB4 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   4
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB4 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   3
            Left            =   1560
            TabIndex        =   39
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB4 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   2
            Left            =   840
            TabIndex        =   38
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB4 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   612
            Index           =   1
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   612
         End
      End
      Begin VB.Frame MainBoard 
         BackColor       =   &H00E0E0E0&
         Height          =   2412
         Index           =   2
         Left            =   2520
         TabIndex        =   11
         Top             =   120
         Width           =   2292
         Begin VB.Label TileB2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   9
            Left            =   1560
            TabIndex        =   20
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   8
            Left            =   840
            TabIndex        =   19
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   7
            Left            =   120
            TabIndex        =   18
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   6
            Left            =   1560
            TabIndex        =   17
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   5
            Left            =   840
            TabIndex        =   16
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   3
            Left            =   1560
            TabIndex        =   14
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   2
            Left            =   840
            TabIndex        =   13
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   612
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   612
         End
      End
      Begin VB.Frame MainBoard 
         BackColor       =   &H00E0E0E0&
         Height          =   2412
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2292
         Begin VB.Label TileB1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   9
            Left            =   1560
            TabIndex        =   10
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   8
            Left            =   840
            TabIndex        =   9
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   7
            Left            =   120
            TabIndex        =   8
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label TileB1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   6
            Left            =   1560
            TabIndex        =   7
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   5
            Left            =   840
            TabIndex        =   6
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   4
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   612
         End
         Begin VB.Label TileB1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   3
            Left            =   1560
            TabIndex        =   4
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   612
            Index           =   2
            Left            =   840
            TabIndex        =   3
            Top             =   240
            Width           =   612
         End
         Begin VB.Label TileB1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   27.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   612
            Index           =   1
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   612
         End
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1560
      TabIndex        =   103
      Top             =   7920
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2760
      TabIndex        =   102
      Top             =   7920
      Width           =   972
   End
   Begin VB.Line Line5 
      X1              =   1320
      X2              =   1320
      Y1              =   7680
      Y2              =   8640
   End
   Begin VB.Line Line4 
      X1              =   7440
      X2              =   7440
      Y1              =   7560
      Y2              =   8760
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   7440
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   7560
      Y2              =   8760
   End
   Begin VB.Label Rep9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   960
      TabIndex        =   101
      Top             =   8400
      Width           =   252
   End
   Begin VB.Label Rep8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   600
      TabIndex        =   100
      Top             =   8400
      Width           =   252
   End
   Begin VB.Label Rep7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   99
      Top             =   8400
      Width           =   252
   End
   Begin VB.Label Rep6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   960
      TabIndex        =   98
      Top             =   8040
      Width           =   252
   End
   Begin VB.Label Rep5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   600
      TabIndex        =   97
      Top             =   8040
      Width           =   252
   End
   Begin VB.Label Rep4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   96
      Top             =   8040
      Width           =   252
   End
   Begin VB.Label Rep3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   960
      TabIndex        =   95
      Top             =   7680
      Width           =   252
   End
   Begin VB.Label Rep2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   600
      TabIndex        =   94
      Top             =   7680
      Width           =   252
   End
   Begin VB.Label Rep1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   93
      Top             =   7680
      Width           =   252
   End
   Begin VB.Label lblNumOfMoves 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6960
      TabIndex        =   92
      Top             =   8040
      Width           =   372
   End
   Begin VB.Label lblNumOfMovesText 
      BackColor       =   &H00C0C0C0&
      Caption         =   "# of Moves:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5760
      TabIndex        =   91
      Top             =   8040
      Width           =   1332
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7440
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Selector 
      BackColor       =   &H00404040&
      Height          =   732
      Left            =   2640
      TabIndex        =   104
      Top             =   7800
      Width           =   1212
   End
End
Attribute VB_Name = "frmT_Cubed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'All Declarations

Dim PlayerTurn As T3Player

Dim BoardVictory(1 To 9) As T3Player

Dim MainBoard_Enabled(1 To 9) As Boolean
Dim TileB1_Played(1 To 9) As Boolean
Dim TileB2_Played(1 To 9) As Boolean
Dim TileB3_Played(1 To 9) As Boolean
Dim TileB4_Played(1 To 9) As Boolean
Dim TileB5_Played(1 To 9) As Boolean
Dim TileB6_Played(1 To 9) As Boolean
Dim TileB7_Played(1 To 9) As Boolean
Dim TileB8_Played(1 To 9) As Boolean
Dim TileB9_Played(1 To 9) As Boolean

Public Enum T3Player
    x
    O
    U
End Enum
'END DECLARATIONS

Private Sub cmdRestart_Click()

    Unload Me
    frmT_CubedMenu.Show

End Sub

Private Sub Form_Load()
    
    PlayerTurn = x
    
    Call StartGame
    
End Sub

Private Sub TileB1_Click(Index As Integer)

    If TileB1_Played(Index) = False Then
        TileB1_Played(Index) = True
        If PlayerTurn = x Then
            TileB1(Index).Caption = "X"
            TileB1(Index).ForeColor = Label1.BackColor
            PlayerTurn = O
        Else
            TileB1(Index).Caption = "O"
            TileB1(Index).ForeColor = Label2.BackColor
            PlayerTurn = x
        End If
        
        Call ChangeBoard(Index)
        Call checkWin(1)
        Call checkDraw(1)
        
    End If
    
End Sub
Private Sub TileB2_Click(Index As Integer)

    If TileB2_Played(Index) = False Then
        TileB2_Played(Index) = True
        If PlayerTurn = x Then
            TileB2(Index).Caption = "X"
            TileB2(Index).ForeColor = Label1.BackColor
            PlayerTurn = O
        Else
            TileB2(Index).Caption = "O"
            TileB2(Index).ForeColor = Label2.BackColor
            PlayerTurn = x
        End If
        
        Call ChangeBoard(Index)
        Call checkWin(2)
        Call checkDraw(2)
        
    End If
    
End Sub
Private Sub TileB3_Click(Index As Integer)

    If TileB3_Played(Index) = False Then
        TileB3_Played(Index) = True
        If PlayerTurn = x Then
            TileB3(Index).Caption = "X"
            TileB3(Index).ForeColor = Label1.BackColor
            PlayerTurn = O
        Else
            TileB3(Index).Caption = "O"
            TileB3(Index).ForeColor = Label2.BackColor
            PlayerTurn = x
        End If
        
        Call ChangeBoard(Index)
        Call checkWin(3)
        Call checkDraw(3)
        
    End If
    
End Sub
Private Sub TileB4_Click(Index As Integer)

    If TileB4_Played(Index) = False Then
        TileB4_Played(Index) = True
        If PlayerTurn = x Then
            TileB4(Index).Caption = "X"
            TileB4(Index).ForeColor = Label1.BackColor
            PlayerTurn = O
        Else
            TileB4(Index).Caption = "O"
            TileB4(Index).ForeColor = Label2.BackColor
            PlayerTurn = x
        End If
        
        Call ChangeBoard(Index)
        Call checkWin(4)
        Call checkDraw(4)
        
    End If
    
End Sub

Private Sub TileB5_Click(Index As Integer)

    If TileB5_Played(Index) = False Then
        TileB5_Played(Index) = True
        
        If PlayerTurn = x Then
            TileB5(Index).Caption = "X"
            TileB5(Index).ForeColor = Label1.BackColor
            PlayerTurn = O
        Else
            TileB5(Index).Caption = "O"
            TileB5(Index).ForeColor = Label2.BackColor
            PlayerTurn = x
        End If
        
        Call ChangeBoard(Index)
        Call checkWin(5)
        Call checkDraw(5)
    End If
    
End Sub
Private Sub TileB6_Click(Index As Integer)

    If TileB6_Played(Index) = False Then
        TileB6_Played(Index) = True
        If PlayerTurn = x Then
            TileB6(Index).Caption = "X"
            TileB6(Index).ForeColor = Label1.BackColor
            PlayerTurn = O
        Else
            TileB6(Index).Caption = "O"
            TileB6(Index).ForeColor = Label2.BackColor
            PlayerTurn = x
        End If
        
        Call ChangeBoard(Index)
        Call checkWin(6)
        Call checkDraw(6)
        
    End If
    
End Sub
Private Sub TileB7_Click(Index As Integer)

    If TileB7_Played(Index) = False Then
        TileB7_Played(Index) = True
        If PlayerTurn = x Then
            TileB7(Index).Caption = "X"
            TileB7(Index).ForeColor = Label1.BackColor
            PlayerTurn = O
        Else
            TileB7(Index).Caption = "O"
            TileB7(Index).ForeColor = Label2.BackColor
            PlayerTurn = x
        End If
        
        Call ChangeBoard(Index)
        Call checkWin(7)
        Call checkDraw(7)
        
    End If
    
End Sub
Private Sub TileB8_Click(Index As Integer)

    If TileB8_Played(Index) = False Then
        TileB8_Played(Index) = True
        If PlayerTurn = x Then
            TileB8(Index).Caption = "X"
            TileB8(Index).ForeColor = Label1.BackColor
            PlayerTurn = O
        Else
            TileB8(Index).Caption = "O"
            TileB8(Index).ForeColor = Label2.BackColor
            PlayerTurn = x
        End If
        
        Call ChangeBoard(Index)
        Call checkWin(8)
        Call checkDraw(8)
        
    End If
    
End Sub

Private Sub TileB9_Click(Index As Integer)

    If TileB9_Played(Index) = False Then
        TileB9_Played(Index) = True
        If PlayerTurn = x Then
            TileB9(Index).Caption = "X"
            TileB9(Index).ForeColor = Label1.BackColor
            PlayerTurn = O
        Else
            TileB9(Index).Caption = "O"
            TileB9(Index).ForeColor = Label2.BackColor
            PlayerTurn = x
        End If
        
        Call ChangeBoard(Index)
        Call checkWin(9)
        Call checkDraw(9)
        
    End If
    
End Sub

Private Sub StartGame()
    Dim i As Integer
    
    For i = 1 To 9
    
        TileB1(i).Caption = ""
        TileB1_Played(i) = False
        TileB2(i).Caption = ""
        TileB2_Played(i) = False
        TileB3(i).Caption = ""
        TileB3_Played(i) = False
        TileB4(i).Caption = ""
        TileB4_Played(i) = False
        TileB5(i).Caption = ""
        TileB5_Played(i) = False
        TileB6(i).Caption = ""
        TileB6_Played(i) = False
        TileB7(i).Caption = ""
        TileB7_Played(i) = False
        TileB8(i).Caption = ""
        TileB8_Played(i) = False
        TileB9(i).Caption = ""
        TileB9_Played(i) = False
        
        BoardVictory(i) = U
        
        If i <> 5 Then
        
            MainBoard_Enabled(i) = False
            MainBoard(i).Enabled = False
            MainBoard(i).BackColor = &HE0E0E0
            
        Else
        
            MainBoard_Enabled(i) = True
            MainBoard_Enabled(i) = True
            MainBoard(i).BackColor = &H808080
            
        End If
        
    Next i

End Sub

Private Sub ChangeBoard(changeToVal As Integer)
    Dim i As Integer
    Dim allOccupied As Integer
    
    lblNumOfMoves = lblNumOfMoves + 1
    
    allOccupied = 0

    If changeToVal = 1 Then
    
        For i = 1 To 9
        
            If TileB1_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 2 Then
    
        For i = 1 To 9
        
            If TileB2_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 3 Then
    
        For i = 1 To 9
        
            If TileB3_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 4 Then
    
        For i = 1 To 9
        
            If TileB4_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 5 Then
    
        For i = 1 To 9
        
            If TileB5_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 6 Then
    
        For i = 1 To 9
        
            If TileB6_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 7 Then
    
        For i = 1 To 9
        
            If TileB7_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 8 Then
    
        For i = 1 To 9
        
            If TileB8_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 9 Then
    
        For i = 1 To 9
        
            If TileB9_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    End If
        
    If allOccupied > 0 Then
    
        For i = 1 To 9
        
            If i <> changeToVal Then
            
                MainBoard_Enabled(i) = False
                MainBoard(i).Enabled = False
                MainBoard(i).BackColor = &HE0E0E0
                MainBoard(i).MousePointer = MousePointerConstants.vbArrow
                
            Else
            
                MainBoard_Enabled(i) = True
                MainBoard(i).Enabled = True
                MainBoard(i).BackColor = &H808080
                MainBoard(i).MousePointer = MousePointerConstants.vbCrosshair
            
            End If
        
        Next i
        
    Else
        
        For i = 1 To 9
            
            If i = changeToVal Then
            
                MainBoard_Enabled(i) = False
                MainBoard(i).Enabled = False
                MainBoard(i).BackColor = &HE0E0E0
                MainBoard(i).MousePointer = MousePointerConstants.vbArrow
                
            Else
            
                MainBoard_Enabled(i) = True
                MainBoard(i).Enabled = True
                MainBoard(i).BackColor = &H808080
                MainBoard(i).MousePointer = MousePointerConstants.vbCrosshair
            
            End If
            
        Next i
        
    End If

End Sub

Private Sub checkWin(BoardNum As Integer)
Dim i As Integer
    Timer1.Enabled = True
    If BoardNum = 1 Then
        
        If BoardVictory(1) = U Then
        
            If (TileB1(1) = "X" And TileB1(2) = "X" And TileB1(3) = "X") Or _
                (TileB1(4) = "X" And TileB1(5) = "X" And TileB1(6) = "X") Or _
                (TileB1(7) = "X" And TileB1(8) = "X" And TileB1(9) = "X") Or _
                (TileB1(1) = "X" And TileB1(4) = "X" And TileB1(7) = "X") Or _
                (TileB1(2) = "X" And TileB1(5) = "X" And TileB1(8) = "X") Or _
                (TileB1(3) = "X" And TileB1(6) = "X" And TileB1(9) = "X") Or _
                (TileB1(1) = "X" And TileB1(5) = "X" And TileB1(9) = "X") Or _
                (TileB1(3) = "X" And TileB1(5) = "X" And TileB1(7) = "X") Then
                
                Rep1.BackColor = Label1.BackColor
                Rep1.Caption = "X"
                For i = 1 To 9
                
                    TileB1(i).ForeColor = Label1.BackColor
                    
                Next i
                
                BoardVictory(1) = x
                
            End If
            
            If (TileB1(1) = "O" And TileB1(2) = "O" And TileB1(3) = "O") Or _
                (TileB1(4) = "O" And TileB1(5) = "O" And TileB1(6) = "O") Or _
                (TileB1(7) = "O" And TileB1(8) = "O" And TileB1(9) = "O") Or _
                (TileB1(1) = "O" And TileB1(4) = "O" And TileB1(7) = "O") Or _
                (TileB1(2) = "O" And TileB1(5) = "O" And TileB1(8) = "O") Or _
                (TileB1(3) = "O" And TileB1(6) = "O" And TileB1(9) = "O") Or _
                (TileB1(1) = "O" And TileB1(5) = "O" And TileB1(9) = "O") Or _
                (TileB1(3) = "O" And TileB1(5) = "O" And TileB1(7) = "O") Then
                
                Rep1.BackColor = Label2.BackColor
                Rep1.Caption = "O"
                For i = 1 To 9
                
                    TileB1(i).ForeColor = Label2.BackColor
                    
                Next i
                
                BoardVictory(1) = O
                
            End If
            
        ElseIf BoardVictory(1) = O Then
        
            Rep1.BackColor = Label2.BackColor
            Rep1.Caption = "O"
            For i = 1 To 9
                
                TileB1(i).ForeColor = Label2.BackColor
                    
            Next i
                
        ElseIf BoardVictory(1) = x Then
        
            Rep1.BackColor = Label1.BackColor
            Rep1.Caption = "X"
            For i = 1 To 9
                
                TileB1(i).ForeColor = Label1.BackColor
                    
            Next i
                
        End If
            
    End If
    
    If BoardNum = 2 Then
        
        If BoardVictory(2) = U Then
        
            If (TileB2(1) = "X" And TileB2(2) = "X" And TileB2(3) = "X") Or _
                (TileB2(4) = "X" And TileB2(5) = "X" And TileB2(6) = "X") Or _
                (TileB2(7) = "X" And TileB2(8) = "X" And TileB2(9) = "X") Or _
                (TileB2(1) = "X" And TileB2(4) = "X" And TileB2(7) = "X") Or _
                (TileB2(2) = "X" And TileB2(5) = "X" And TileB2(8) = "X") Or _
                (TileB2(3) = "X" And TileB2(6) = "X" And TileB2(9) = "X") Or _
                (TileB2(1) = "X" And TileB2(5) = "X" And TileB2(9) = "X") Or _
                (TileB2(3) = "X" And TileB2(5) = "X" And TileB2(7) = "X") Then
                
                Rep2.BackColor = Label1.BackColor
                Rep2.Caption = "X"
                For i = 1 To 9
                
                    TileB2(i).ForeColor = Label1.BackColor
                    
                Next i
                
                BoardVictory(2) = x
                
            End If
            
            If (TileB2(1) = "O" And TileB2(2) = "O" And TileB2(3) = "O") Or _
                (TileB2(4) = "O" And TileB2(5) = "O" And TileB2(6) = "O") Or _
                (TileB2(7) = "O" And TileB2(8) = "O" And TileB2(9) = "O") Or _
                (TileB2(1) = "O" And TileB2(4) = "O" And TileB2(7) = "O") Or _
                (TileB2(2) = "O" And TileB2(5) = "O" And TileB2(8) = "O") Or _
                (TileB2(3) = "O" And TileB2(6) = "O" And TileB2(9) = "O") Or _
                (TileB2(1) = "O" And TileB2(5) = "O" And TileB2(9) = "O") Or _
                (TileB2(3) = "O" And TileB2(5) = "O" And TileB2(7) = "O") Then
                
                Rep2.BackColor = Label2.BackColor
                Rep2.Caption = "O"
                For i = 1 To 9
                
                    TileB2(i).ForeColor = Label2.BackColor
                    
                Next i
                
                BoardVictory(2) = O
                
            End If
            
        ElseIf BoardVictory(2) = O Then
        
            Rep2.BackColor = Label2.BackColor
            Rep2.Caption = "O"
            For i = 1 To 9
                
                TileB2(i).ForeColor = Label2.BackColor
                    
            Next i
                
        ElseIf BoardVictory(2) = x Then
        
            Rep2.BackColor = Label1.BackColor
            Rep2.Caption = "X"
            For i = 1 To 9
                
                TileB2(i).ForeColor = Label1.BackColor
                    
            Next i
                
        End If
            
    End If

    If BoardNum = 3 Then
        
        If BoardVictory(3) = U Then
        
            If (TileB3(1) = "X" And TileB3(2) = "X" And TileB3(3) = "X") Or _
                (TileB3(4) = "X" And TileB3(5) = "X" And TileB3(6) = "X") Or _
                (TileB3(7) = "X" And TileB3(8) = "X" And TileB3(9) = "X") Or _
                (TileB3(1) = "X" And TileB3(4) = "X" And TileB3(7) = "X") Or _
                (TileB3(2) = "X" And TileB3(5) = "X" And TileB3(8) = "X") Or _
                (TileB3(3) = "X" And TileB3(6) = "X" And TileB3(9) = "X") Or _
                (TileB3(1) = "X" And TileB3(5) = "X" And TileB3(9) = "X") Or _
                (TileB3(3) = "X" And TileB3(5) = "X" And TileB3(7) = "X") Then
                
                Rep3.BackColor = Label1.BackColor
                Rep3.Caption = "X"
                For i = 1 To 9
                
                    TileB3(i).ForeColor = Label1.BackColor
                    
                Next i
                
                BoardVictory(3) = x
                
            End If
            
            If (TileB3(1) = "O" And TileB3(2) = "O" And TileB3(3) = "O") Or _
                (TileB3(4) = "O" And TileB3(5) = "O" And TileB3(6) = "O") Or _
                (TileB3(7) = "O" And TileB3(8) = "O" And TileB3(9) = "O") Or _
                (TileB3(1) = "O" And TileB3(4) = "O" And TileB3(7) = "O") Or _
                (TileB3(2) = "O" And TileB3(5) = "O" And TileB3(8) = "O") Or _
                (TileB3(3) = "O" And TileB3(6) = "O" And TileB3(9) = "O") Or _
                (TileB3(1) = "O" And TileB3(5) = "O" And TileB3(9) = "O") Or _
                (TileB3(3) = "O" And TileB3(5) = "O" And TileB3(7) = "O") Then
                
                Rep3.BackColor = Label2.BackColor
                Rep3.Caption = "O"
                For i = 1 To 9
                
                    TileB3(i).ForeColor = Label2.BackColor
                    
                Next i
                
                BoardVictory(3) = O
                
            End If
            
        ElseIf BoardVictory(3) = O Then
        
            Rep3.BackColor = Label2.BackColor
            Rep3.Caption = "O"
            For i = 1 To 9
                
                TileB3(i).ForeColor = Label2.BackColor
                    
            Next i
                
        ElseIf BoardVictory(3) = x Then
        
            Rep3.BackColor = Label1.BackColor
            Rep3.Caption = "X"
            For i = 1 To 9
                
                TileB3(i).ForeColor = Label1.BackColor
                    
            Next i
                
        End If
            
    End If
    
    If BoardNum = 4 Then
        
        If BoardVictory(4) = U Then
        
            If (TileB4(1) = "X" And TileB4(2) = "X" And TileB4(3) = "X") Or _
                (TileB4(4) = "X" And TileB4(5) = "X" And TileB4(6) = "X") Or _
                (TileB4(7) = "X" And TileB4(8) = "X" And TileB4(9) = "X") Or _
                (TileB4(1) = "X" And TileB4(4) = "X" And TileB4(7) = "X") Or _
                (TileB4(2) = "X" And TileB4(5) = "X" And TileB4(8) = "X") Or _
                (TileB4(3) = "X" And TileB4(6) = "X" And TileB4(9) = "X") Or _
                (TileB4(1) = "X" And TileB4(5) = "X" And TileB4(9) = "X") Or _
                (TileB4(3) = "X" And TileB4(5) = "X" And TileB4(7) = "X") Then
                
                Rep4.BackColor = Label1.BackColor
                Rep4.Caption = "X"
                For i = 1 To 9
                
                    TileB4(i).ForeColor = Label1.BackColor
                    
                Next i
                
                BoardVictory(4) = x
                
            End If
            
            If (TileB4(1) = "O" And TileB4(2) = "O" And TileB4(3) = "O") Or _
                (TileB4(4) = "O" And TileB4(5) = "O" And TileB4(6) = "O") Or _
                (TileB4(7) = "O" And TileB4(8) = "O" And TileB4(9) = "O") Or _
                (TileB4(1) = "O" And TileB4(4) = "O" And TileB4(7) = "O") Or _
                (TileB4(2) = "O" And TileB4(5) = "O" And TileB4(8) = "O") Or _
                (TileB4(3) = "O" And TileB4(6) = "O" And TileB4(9) = "O") Or _
                (TileB4(1) = "O" And TileB4(5) = "O" And TileB4(9) = "O") Or _
                (TileB4(3) = "O" And TileB4(5) = "O" And TileB4(7) = "O") Then
                
                Rep4.BackColor = Label2.BackColor
                Rep4.Caption = "O"
                For i = 1 To 9
                
                    TileB4(i).ForeColor = Label2.BackColor
                    
                Next i
                
                BoardVictory(4) = O
                
            End If
            
        ElseIf BoardVictory(4) = O Then
        
            Rep4.BackColor = Label2.BackColor
            Rep4.Caption = "O"
            For i = 1 To 9
                
                TileB4(i).ForeColor = Label2.BackColor
                    
            Next i
                
        ElseIf BoardVictory(4) = x Then
        
            Rep4.BackColor = Label1.BackColor
            Rep4.Caption = "X"
            For i = 1 To 9
                
                TileB4(i).ForeColor = Label1.BackColor
                    
            Next i
                
        End If
            
    End If

    If BoardNum = 5 Then
        
        If BoardVictory(5) = U Then
        
            If (TileB5(1) = "X" And TileB5(2) = "X" And TileB5(3) = "X") Or _
                (TileB5(4) = "X" And TileB5(5) = "X" And TileB5(6) = "X") Or _
                (TileB5(7) = "X" And TileB5(8) = "X" And TileB5(9) = "X") Or _
                (TileB5(1) = "X" And TileB5(4) = "X" And TileB5(7) = "X") Or _
                (TileB5(2) = "X" And TileB5(5) = "X" And TileB5(8) = "X") Or _
                (TileB5(3) = "X" And TileB5(6) = "X" And TileB5(9) = "X") Or _
                (TileB5(1) = "X" And TileB5(5) = "X" And TileB5(9) = "X") Or _
                (TileB5(3) = "X" And TileB5(5) = "X" And TileB5(7) = "X") Then
                
                Rep5.BackColor = Label1.BackColor
                Rep5.Caption = "X"
                For i = 1 To 9
                
                    TileB5(i).ForeColor = Label1.BackColor
                    
                Next i
                
                BoardVictory(5) = x
                
            End If
            
            If (TileB5(1) = "O" And TileB5(2) = "O" And TileB5(3) = "O") Or _
                (TileB5(4) = "O" And TileB5(5) = "O" And TileB5(6) = "O") Or _
                (TileB5(7) = "O" And TileB5(8) = "O" And TileB5(9) = "O") Or _
                (TileB5(1) = "O" And TileB5(4) = "O" And TileB5(7) = "O") Or _
                (TileB5(2) = "O" And TileB5(5) = "O" And TileB5(8) = "O") Or _
                (TileB5(3) = "O" And TileB5(6) = "O" And TileB5(9) = "O") Or _
                (TileB5(1) = "O" And TileB5(5) = "O" And TileB5(9) = "O") Or _
                (TileB5(3) = "O" And TileB5(5) = "O" And TileB5(7) = "O") Then
                
                Rep5.BackColor = Label2.BackColor
                Rep5.Caption = "O"
                For i = 1 To 9
                
                    TileB5(i).ForeColor = Label2.BackColor
                    
                Next i
                
                BoardVictory(5) = O
                
            End If
            
        ElseIf BoardVictory(5) = O Then
        
            Rep5.BackColor = Label2.BackColor
            Rep5.Caption = "O"
            For i = 1 To 9
                
                TileB5(i).ForeColor = Label2.BackColor
                    
            Next i
                
        ElseIf BoardVictory(5) = x Then
        
            Rep5.BackColor = Label1.BackColor
            Rep5.Caption = "X"
            For i = 1 To 9
                
                TileB5(i).ForeColor = Label1.BackColor
                    
            Next i
                
        End If
              
    End If
    
    If BoardNum = 6 Then
        
        If BoardVictory(6) = U Then
        
            If (TileB6(1) = "X" And TileB6(2) = "X" And TileB6(3) = "X") Or _
                (TileB6(4) = "X" And TileB6(5) = "X" And TileB6(6) = "X") Or _
                (TileB6(7) = "X" And TileB6(8) = "X" And TileB6(9) = "X") Or _
                (TileB6(1) = "X" And TileB6(4) = "X" And TileB6(7) = "X") Or _
                (TileB6(2) = "X" And TileB6(5) = "X" And TileB6(8) = "X") Or _
                (TileB6(3) = "X" And TileB6(6) = "X" And TileB6(9) = "X") Or _
                (TileB6(1) = "X" And TileB6(5) = "X" And TileB6(9) = "X") Or _
                (TileB6(3) = "X" And TileB6(5) = "X" And TileB6(7) = "X") Then
                
                Rep6.BackColor = Label1.BackColor
                Rep6.Caption = "X"
                For i = 1 To 9
                
                    TileB6(i).ForeColor = Label1.BackColor
                    
                Next i
                
                BoardVictory(6) = x
                
            End If
            
            If (TileB6(1) = "O" And TileB6(2) = "O" And TileB6(3) = "O") Or _
                (TileB6(4) = "O" And TileB6(5) = "O" And TileB6(6) = "O") Or _
                (TileB6(7) = "O" And TileB6(8) = "O" And TileB6(9) = "O") Or _
                (TileB6(1) = "O" And TileB6(4) = "O" And TileB6(7) = "O") Or _
                (TileB6(2) = "O" And TileB6(5) = "O" And TileB6(8) = "O") Or _
                (TileB6(3) = "O" And TileB6(6) = "O" And TileB6(9) = "O") Or _
                (TileB6(1) = "O" And TileB6(5) = "O" And TileB6(9) = "O") Or _
                (TileB6(3) = "O" And TileB6(5) = "O" And TileB6(7) = "O") Then
                
                Rep6.BackColor = Label2.BackColor
                Rep6.Caption = "O"
                For i = 1 To 9
                
                    TileB6(i).ForeColor = Label2.BackColor
                    
                Next i
                
                BoardVictory(6) = O
                
            End If
            
        ElseIf BoardVictory(6) = O Then
        
            Rep6.BackColor = Label2.BackColor
            Rep6.Caption = "O"
            For i = 1 To 9
                
                TileB6(i).ForeColor = Label2.BackColor
                    
            Next i
                
        ElseIf BoardVictory(6) = x Then
        
            Rep6.BackColor = Label1.BackColor
            Rep6.Caption = "X"
            For i = 1 To 9
                
                TileB6(i).ForeColor = Label1.BackColor
                    
            Next i
                
        End If
            
    End If
    
    If BoardNum = 7 Then
        
        If BoardVictory(7) = U Then
        
            If (TileB7(1) = "X" And TileB7(2) = "X" And TileB7(3) = "X") Or _
                (TileB7(4) = "X" And TileB7(5) = "X" And TileB7(6) = "X") Or _
                (TileB7(7) = "X" And TileB7(8) = "X" And TileB7(9) = "X") Or _
                (TileB7(1) = "X" And TileB7(4) = "X" And TileB7(7) = "X") Or _
                (TileB7(2) = "X" And TileB7(5) = "X" And TileB7(8) = "X") Or _
                (TileB7(3) = "X" And TileB7(6) = "X" And TileB7(9) = "X") Or _
                (TileB7(1) = "X" And TileB7(5) = "X" And TileB7(9) = "X") Or _
                (TileB7(3) = "X" And TileB7(5) = "X" And TileB7(7) = "X") Then
                
                Rep7.BackColor = Label1.BackColor
                Rep7.Caption = "X"
                For i = 1 To 9
                
                    TileB7(i).ForeColor = Label1.BackColor
                    
                Next i
                
                BoardVictory(7) = x
                
            End If
            
            If (TileB7(1) = "O" And TileB7(2) = "O" And TileB7(3) = "O") Or _
                (TileB7(4) = "O" And TileB7(5) = "O" And TileB7(6) = "O") Or _
                (TileB7(7) = "O" And TileB7(8) = "O" And TileB7(9) = "O") Or _
                (TileB7(1) = "O" And TileB7(4) = "O" And TileB7(7) = "O") Or _
                (TileB7(2) = "O" And TileB7(5) = "O" And TileB7(8) = "O") Or _
                (TileB7(3) = "O" And TileB7(6) = "O" And TileB7(9) = "O") Or _
                (TileB7(1) = "O" And TileB7(5) = "O" And TileB7(9) = "O") Or _
                (TileB7(3) = "O" And TileB7(5) = "O" And TileB7(7) = "O") Then
                
                Rep7.BackColor = Label2.BackColor
                Rep7.Caption = "O"
                For i = 1 To 9
                
                    TileB7(i).ForeColor = Label2.BackColor
                    
                Next i
                
                BoardVictory(7) = O
                
            End If
            
        ElseIf BoardVictory(7) = O Then
        
            Rep7.BackColor = Label2.BackColor
            Rep7.Caption = "O"
            For i = 1 To 9
                
                TileB7(i).ForeColor = Label2.BackColor
                    
            Next i
                
        ElseIf BoardVictory(7) = x Then
        
            Rep7.BackColor = Label1.BackColor
            Rep7.Caption = "X"
            For i = 1 To 9
                
                TileB7(i).ForeColor = Label1.BackColor
                    
            Next i
                
        End If
            
    End If
    
    If BoardNum = 8 Then
        
        If BoardVictory(8) = U Then
        
            If (TileB8(1) = "X" And TileB8(2) = "X" And TileB8(3) = "X") Or _
                (TileB8(4) = "X" And TileB8(5) = "X" And TileB8(6) = "X") Or _
                (TileB8(7) = "X" And TileB8(8) = "X" And TileB8(9) = "X") Or _
                (TileB8(1) = "X" And TileB8(4) = "X" And TileB8(7) = "X") Or _
                (TileB8(2) = "X" And TileB8(5) = "X" And TileB8(8) = "X") Or _
                (TileB8(3) = "X" And TileB8(6) = "X" And TileB8(9) = "X") Or _
                (TileB8(1) = "X" And TileB8(5) = "X" And TileB8(9) = "X") Or _
                (TileB8(3) = "X" And TileB8(5) = "X" And TileB8(7) = "X") Then
                
                Rep8.BackColor = Label1.BackColor
                Rep8.Caption = "X"
                For i = 1 To 9
                
                    TileB8(i).ForeColor = Label1.BackColor
                    
                Next i
                
                BoardVictory(8) = x
                
            End If
            
            If (TileB8(1) = "O" And TileB8(2) = "O" And TileB8(3) = "O") Or _
                (TileB8(4) = "O" And TileB8(5) = "O" And TileB8(6) = "O") Or _
                (TileB8(7) = "O" And TileB8(8) = "O" And TileB8(9) = "O") Or _
                (TileB8(1) = "O" And TileB8(4) = "O" And TileB8(7) = "O") Or _
                (TileB8(2) = "O" And TileB8(5) = "O" And TileB8(8) = "O") Or _
                (TileB8(3) = "O" And TileB8(6) = "O" And TileB8(9) = "O") Or _
                (TileB8(1) = "O" And TileB8(5) = "O" And TileB8(9) = "O") Or _
                (TileB8(3) = "O" And TileB8(5) = "O" And TileB8(7) = "O") Then
                
                Rep8.BackColor = Label2.BackColor
                Rep8.Caption = "O"
                For i = 1 To 9
                
                    TileB8(i).ForeColor = Label2.BackColor
                    
                Next i
                
                BoardVictory(8) = O
                
            End If
            
        ElseIf BoardVictory(8) = O Then
        
            Rep8.BackColor = Label2.BackColor
            Rep8.Caption = "O"
            For i = 1 To 9
                
                TileB8(i).ForeColor = Label2.BackColor
                    
            Next i
                
        ElseIf BoardVictory(8) = x Then
        
            Rep8.BackColor = Label1.BackColor
            Rep8.Caption = "X"
            For i = 1 To 9
                
                TileB8(i).ForeColor = Label1.BackColor
                    
            Next i
                
        End If
            
    End If
    
    If BoardNum = 9 Then
        
        If BoardVictory(9) = U Then
        
            If (TileB9(1) = "X" And TileB9(2) = "X" And TileB9(3) = "X") Or _
                (TileB9(4) = "X" And TileB9(5) = "X" And TileB9(6) = "X") Or _
                (TileB9(7) = "X" And TileB9(8) = "X" And TileB9(9) = "X") Or _
                (TileB9(1) = "X" And TileB9(4) = "X" And TileB9(7) = "X") Or _
                (TileB9(2) = "X" And TileB9(5) = "X" And TileB9(8) = "X") Or _
                (TileB9(3) = "X" And TileB9(6) = "X" And TileB9(9) = "X") Or _
                (TileB9(1) = "X" And TileB9(5) = "X" And TileB9(9) = "X") Or _
                (TileB9(3) = "X" And TileB9(5) = "X" And TileB9(7) = "X") Then
                
                Rep9.BackColor = Label1.BackColor
                Rep9.Caption = "X"
                For i = 1 To 9
                
                    TileB9(i).ForeColor = Label1.BackColor
                    
                Next i
                
                BoardVictory(9) = x
                
            End If
            
            If (TileB9(1) = "O" And TileB9(2) = "O" And TileB9(3) = "O") Or _
                (TileB9(4) = "O" And TileB9(5) = "O" And TileB9(6) = "O") Or _
                (TileB9(7) = "O" And TileB9(8) = "O" And TileB9(9) = "O") Or _
                (TileB9(1) = "O" And TileB9(4) = "O" And TileB9(7) = "O") Or _
                (TileB9(2) = "O" And TileB9(5) = "O" And TileB9(8) = "O") Or _
                (TileB9(3) = "O" And TileB9(6) = "O" And TileB9(9) = "O") Or _
                (TileB9(1) = "O" And TileB9(5) = "O" And TileB9(9) = "O") Or _
                (TileB9(3) = "O" And TileB9(5) = "O" And TileB9(7) = "O") Then
                
                Rep9.BackColor = Label2.BackColor
                Rep9.Caption = "O"
                For i = 1 To 9
                
                    TileB9(i).ForeColor = Label2.BackColor
                    
                Next i
                
                BoardVictory(9) = O
                
            End If
            
        ElseIf BoardVictory(9) = O Then
        
            Rep9.BackColor = Label2.BackColor
            Rep9.Caption = "O"
            For i = 1 To 9
                
                TileB9(i).ForeColor = Label2.BackColor
                    
            Next i
                
        ElseIf BoardVictory(9) = x Then
        
            Rep9.BackColor = Label1.BackColor
            Rep9.Caption = "X"
            For i = 1 To 9
                
                TileB9(i).ForeColor = Label1.BackColor
                    
            Next i
                
        End If
 
    End If
    
    Call hasHeWon(x)
    Call hasHeWon(O)

End Sub

Private Sub checkDraw(changeToVal As Integer)
    Dim i As Integer
    Dim allOccupied As Integer
    
    allOccupied = 0

    If changeToVal = 1 Then
    
        For i = 1 To 9
        
            If TileB1_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 2 Then
    
        For i = 1 To 9
        
            If TileB2_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 3 Then
    
        For i = 1 To 9
        
            If TileB3_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 4 Then
    
        For i = 1 To 9
        
            If TileB4_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 5 Then
    
        For i = 1 To 9
        
            If TileB5_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 6 Then
    
        For i = 1 To 9
        
            If TileB6_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 7 Then
    
        For i = 1 To 9
        
            If TileB7_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 8 Then
    
        For i = 1 To 9
        
            If TileB8_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    ElseIf changeToVal = 9 Then
    
        For i = 1 To 9
        
            If TileB9_Played(i) = False Then
            
                allOccupied = allOccupied + 1
                Exit For
                
            End If
            
        Next i
        
    End If

    If allOccupied = 0 Then

        If BoardVictory(changeToVal) = U Then

            If changeToVal = 1 Then
            
                Rep1.Caption = "D"
                Rep1.BackColor = ColorConstants.vbWhite
                
            ElseIf changeToVal = 2 Then
            
                Rep2.Caption = "D"
                Rep2.BackColor = ColorConstants.vbWhite
                
            ElseIf changeToVal = 3 Then
            
                Rep3.Caption = "D"
                Rep3.BackColor = ColorConstants.vbWhite
                
            ElseIf changeToVal = 4 Then
            
                Rep4.Caption = "D"
                Rep4.BackColor = ColorConstants.vbWhite
                
            ElseIf changeToVal = 5 Then
            
                Rep5.Caption = "D"
                Rep5.BackColor = ColorConstants.vbWhite
                
            ElseIf changeToVal = 6 Then
            
                Rep6.Caption = "D"
                Rep6.BackColor = ColorConstants.vbWhite
                
            ElseIf changeToVal = 7 Then
            
                Rep7.Caption = "D"
                Rep7.BackColor = ColorConstants.vbWhite
                
            ElseIf changeToVal = 8 Then
            
                Rep8.Caption = "D"
                Rep8.BackColor = ColorConstants.vbWhite
                
            ElseIf changeToVal = 9 Then
            
                Rep9.Caption = "D"
                Rep9.BackColor = ColorConstants.vbWhite
                
            End If
            
        End If
        
    End If
    
    Call hasHeWon(x)
    Call hasHeWon(O)
        
End Sub

Private Sub Timer1_Timer()
    
    If PlayerTurn = O Then
    
        If Selector.Left <> 1440 Then
        
            Selector.Left = Selector.Left - 120
            
        Else
        
            Timer1.Enabled = False
        
        End If
        
    ElseIf PlayerTurn = x Then
    
        If Selector.Left <> 2640 Then
    
            Selector.Left = Selector.Left + 120
            
        Else
        
            Timer1.Enabled = False
        
        End If
        
    End If
    
End Sub

Private Sub hasHeWon(Player As T3Player)

    If Player = x Then
    
        If (Rep1 = "X" And Rep2 = "X" And Rep3 = "X") Or _
            (Rep1 = "X" And Rep2 = "X" And Rep3 = "D") Or _
            (Rep1 = "X" And Rep2 = "D" And Rep3 = "X") Or _
            (Rep1 = "D" And Rep2 = "X" And Rep3 = "X") Or _
            (Rep4 = "X" And Rep5 = "X" And Rep6 = "X") Or _
            (Rep4 = "X" And Rep5 = "X" And Rep6 = "D") Or _
            (Rep4 = "X" And Rep5 = "D" And Rep6 = "X") Or _
            (Rep4 = "D" And Rep5 = "X" And Rep6 = "X") Or _
            (Rep7 = "X" And Rep8 = "X" And Rep9 = "X") Or _
            (Rep7 = "X" And Rep8 = "X" And Rep9 = "D") Or _
            (Rep7 = "X" And Rep8 = "D" And Rep9 = "X") Or _
            (Rep7 = "D" And Rep8 = "X" And Rep9 = "X") Or _
            (Rep1 = "X" And Rep4 = "X" And Rep7 = "X") Or _
            (Rep1 = "X" And Rep4 = "X" And Rep7 = "D") Or _
            (Rep1 = "X" And Rep4 = "D" And Rep7 = "X") Or _
            (Rep1 = "D" And Rep4 = "X" And Rep7 = "X") Or _
            (Rep2 = "X" And Rep5 = "X" And Rep8 = "X") Or _
            (Rep2 = "X" And Rep5 = "X" And Rep8 = "D") Or _
            (Rep2 = "X" And Rep5 = "D" And Rep8 = "X") Or _
            (Rep2 = "D" And Rep5 = "X" And Rep8 = "X") Or _
            (Rep3 = "X" And Rep6 = "X" And Rep9 = "X") Or _
            (Rep3 = "X" And Rep6 = "X" And Rep9 = "D") Or _
            (Rep3 = "X" And Rep6 = "D" And Rep9 = "X") Or _
            (Rep3 = "D" And Rep6 = "X" And Rep9 = "X") Then
            
                MsgBox "PLAYER X WINS!"
                PlayerTurn = U
            
        End If
        
        If (Rep1 = "X" And Rep5 = "X" And Rep9 = "X") Or _
            (Rep1 = "X" And Rep5 = "X" And Rep9 = "D") Or _
            (Rep1 = "X" And Rep5 = "D" And Rep9 = "X") Or _
            (Rep1 = "D" And Rep5 = "X" And Rep9 = "X") Or _
            (Rep3 = "X" And Rep5 = "X" And Rep7 = "X") Or _
            (Rep3 = "X" And Rep5 = "X" And Rep7 = "D") Or _
            (Rep3 = "X" And Rep5 = "D" And Rep7 = "X") Or _
            (Rep3 = "D" And Rep5 = "X" And Rep7 = "X") Then
                
                MsgBox "PLAYER X WINS!"
                PlayerTurn = U
                
        End If

    End If
    
        If Player = O Then
    
        If (Rep1 = "O" And Rep2 = "O" And Rep3 = "O") Or _
            (Rep1 = "O" And Rep2 = "O" And Rep3 = "D") Or _
            (Rep1 = "O" And Rep2 = "D" And Rep3 = "O") Or _
            (Rep1 = "D" And Rep2 = "O" And Rep3 = "O") Or _
            (Rep4 = "O" And Rep5 = "O" And Rep6 = "O") Or _
            (Rep4 = "O" And Rep5 = "O" And Rep6 = "D") Or _
            (Rep4 = "O" And Rep5 = "D" And Rep6 = "O") Or _
            (Rep4 = "D" And Rep5 = "O" And Rep6 = "O") Or _
            (Rep7 = "O" And Rep8 = "O" And Rep9 = "O") Or _
            (Rep7 = "O" And Rep8 = "O" And Rep9 = "D") Or _
            (Rep7 = "O" And Rep8 = "D" And Rep9 = "O") Or _
            (Rep7 = "D" And Rep8 = "O" And Rep9 = "O") Or _
            (Rep1 = "O" And Rep4 = "O" And Rep7 = "O") Or _
            (Rep1 = "O" And Rep4 = "O" And Rep7 = "D") Or _
            (Rep1 = "O" And Rep4 = "D" And Rep7 = "O") Or _
            (Rep1 = "D" And Rep4 = "O" And Rep7 = "O") Or _
            (Rep2 = "O" And Rep5 = "O" And Rep8 = "O") Or _
            (Rep2 = "O" And Rep5 = "O" And Rep8 = "D") Or _
            (Rep2 = "O" And Rep5 = "D" And Rep8 = "O") Or _
            (Rep2 = "D" And Rep5 = "O" And Rep8 = "O") Or _
            (Rep3 = "O" And Rep6 = "O" And Rep9 = "O") Or _
            (Rep3 = "O" And Rep6 = "O" And Rep9 = "D") Or _
            (Rep3 = "O" And Rep6 = "D" And Rep9 = "O") Or _
            (Rep3 = "D" And Rep6 = "O" And Rep9 = "O") Then
            
                MsgBox "PLAYER O WINS!"
                PlayerTurn = U
            
        End If
        
        If (Rep1 = "O" And Rep5 = "O" And Rep9 = "O") Or _
            (Rep1 = "O" And Rep5 = "O" And Rep9 = "D") Or _
            (Rep1 = "O" And Rep5 = "D" And Rep9 = "O") Or _
            (Rep1 = "D" And Rep5 = "O" And Rep9 = "O") Or _
            (Rep3 = "O" And Rep5 = "O" And Rep7 = "O") Or _
            (Rep3 = "O" And Rep5 = "O" And Rep7 = "D") Or _
            (Rep3 = "O" And Rep5 = "D" And Rep7 = "O") Or _
            (Rep3 = "D" And Rep5 = "O" And Rep7 = "O") Then
                
                MsgBox "PLAYER O WINS!"
                PlayerTurn = U
                
        End If

    End If

End Sub
