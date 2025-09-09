VERSION 5.00
Begin VB.Form About 
   Caption         =   "About WRPN"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   2033
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Dated: 10 Mar 99"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Emmet P. Gray                graye@hood-emh3.army.mil"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Version: 2.0"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   240
      Picture         =   "about.frx":0000
      Top             =   720
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "WRPN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload About
End Sub

