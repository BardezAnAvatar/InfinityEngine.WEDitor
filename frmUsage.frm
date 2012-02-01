VERSION 5.00
Begin VB.Form frmUsage 
   Caption         =   "Usage"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Complex? Yes. This is why I suggest you make sure you know what you are doing."
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   7680
      Width           =   4215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   $"frmUsage.frx":0000
      Height          =   1935
      Left            =   240
      TabIndex        =   3
      Top             =   5640
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   $"frmUsage.frx":021B
      Height          =   4335
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "As stated in Help > About, I hightly recommend that you read the WED file description, so that you know what you are doing."
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "USAGE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
