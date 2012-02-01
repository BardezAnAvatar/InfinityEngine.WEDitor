VERSION 5.00
Begin VB.Form frmPolygonFlag 
   Caption         =   "Also used in game areas, in many places, but at first glance it doesn't seem to do anything. "
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   5115
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1980
      TabIndex        =   20
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "unknown, I have seen this in doors walls flags.  "
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   6000
      Width           =   3375
   End
   Begin VB.Label Label19 
      Caption         =   "unknown, unused ?  "
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "unknown, unused ?  "
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "unknown, unused ?  "
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "Also used in game areas, in many places, but at first glance it doesn't seem to do anything. "
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   600
      TabIndex        =   15
      Top             =   4320
      Width           =   4335
   End
   Begin VB.Label Label15 
      Caption         =   $"frmPolygonFlag.frx":0000
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   600
      TabIndex        =   14
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Label Label14 
      Caption         =   "If set to 1 this disable wall appearance at all. "
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label Label13 
      Caption         =   "Here we have 8 bit flag :"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "The default is that wall shades animations from both sides - it is modified by flag."
      Height          =   495
      Left            =   600
      TabIndex        =   11
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label11 
      Caption         =   "Taken Directly From IESDP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label10 
      Caption         =   "Indicates whether this polygon is a polygon or a hole (i.e. whether it is a section of impassability, or one of passability.) "
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label9 
      Caption         =   $"frmPolygonFlag.frx":00D2
      Height          =   975
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label8 
      Caption         =   "Bit 7:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Bit 6:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Bit 5:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Bit 4:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Bit 3:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Bit 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Bit 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Bit 0:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   375
   End
End
Attribute VB_Name = "frmPolygonFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub
