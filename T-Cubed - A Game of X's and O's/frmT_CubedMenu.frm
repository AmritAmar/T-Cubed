VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmT_CubedMenu 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T-Cubed - A Game of X's and O's"
   ClientHeight    =   4560
   ClientLeft      =   7320
   ClientTop       =   3960
   ClientWidth     =   7632
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7632
   Begin VB.CommandButton cmdStart 
      Caption         =   "Begin Game!"
      Height          =   372
      Left            =   5760
      TabIndex        =   6
      Top             =   3960
      Width           =   1692
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdChangeY 
      Caption         =   "Change O Colour"
      Height          =   372
      Left            =   3720
      TabIndex        =   5
      Top             =   3960
      Width           =   1692
   End
   Begin VB.CommandButton cmdChangeX 
      Caption         =   "Change X Colour"
      Height          =   372
      Left            =   960
      TabIndex        =   4
      Top             =   3960
      Width           =   1692
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2892
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmT_CubedMenu.frx":0000
      Top             =   840
      Width           =   7452
   End
   Begin VB.Label TileO 
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
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Top             =   3840
      Width           =   612
   End
   Begin VB.Label TileX 
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
      Top             =   3840
      Width           =   612
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "T-Cubed"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   25.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7452
   End
End
Attribute VB_Name = "frmT_CubedMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChangeX_Click()

    CommonDialog1.ShowColor
    TileX(1).ForeColor = CommonDialog1.Color

End Sub

Private Sub cmdChangeY_Click()

    CommonDialog1.ShowColor
    TileO(0).ForeColor = CommonDialog1.Color

End Sub

Private Sub cmdStart_Click()

    MsgBox "Thank You for playing the game! Coded and Designed by SirHack3r."
    frmT_Cubed.Show
    frmT_Cubed.Label1.BackColor = frmT_CubedMenu.TileX(1).ForeColor
    frmT_Cubed.Label2.BackColor = frmT_CubedMenu.TileO(0).ForeColor
    Unload Me

End Sub
