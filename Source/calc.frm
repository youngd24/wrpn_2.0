VERSION 5.00
Begin VB.Form Calc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reverse Polish Notation Calculator"
   ClientHeight    =   3720
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6600
   ForeColor       =   &H80000008&
   Icon            =   "calc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   6120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   113
      Top             =   0
      Width           =   1200
   End
   Begin VB.TextBox Annunciator 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   720
      TabIndex        =   112
      Top             =   545
      Width           =   4095
   End
   Begin VB.CommandButton BN_PLUS 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   39
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton BN_CHS 
      Caption         =   "CHS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   38
      ToolTipText     =   "Change sign"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton BN_DP 
      Caption         =   "·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   37
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton BN_0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   36
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton BN_ENTER 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   35
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton BN_RCL 
      Caption         =   "RCL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   34
      ToolTipText     =   "Recall"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton BN_STO 
      Caption         =   "STO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   33
      ToolTipText     =   "Store"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton BN_GKEY 
      BackColor       =   &H00FFFF00&
      Caption         =   "alt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Blue (alt key)"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton BN_FKEY 
      BackColor       =   &H0000FFFF&
      Caption         =   "ctrl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Yellow (ctrl key)"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton BN_ON 
      Caption         =   "ON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   30
      ToolTipText     =   "On/Off"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton BN_MINUS 
      Caption         =   "-"
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
      Left            =   5760
      TabIndex        =   29
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton BN_3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   28
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton BN_2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   27
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton BN_1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   26
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton BN_BSP 
      Caption         =   "BSP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   25
      ToolTipText     =   "Backspace"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton BN_XY 
      Caption         =   "X:Y"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   24
      ToolTipText     =   "Exchange x y"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton BN_ROL 
      Caption         =   "ROL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   23
      ToolTipText     =   "Roll down stack"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton BN_SST 
      Caption         =   "SST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   22
      ToolTipText     =   "Single step"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton BN_RS 
      Caption         =   "R/S"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   21
      ToolTipText     =   "Run/Stop"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton BN_MULT 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   20
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton BN_6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   19
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton BN_5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton BN_4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton BN_BIN 
      Caption         =   "BIN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      ToolTipText     =   "Binary mode"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton BN_OCT 
      Caption         =   "OCT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      ToolTipText     =   "Octal mode"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton BN_DEC 
      Caption         =   "DEC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      ToolTipText     =   "Decimal mode"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton BN_HEX 
      Caption         =   "HEX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      ToolTipText     =   "Hexadecimal mode"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton BN_GTO 
      Caption         =   "GTO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   12
      ToolTipText     =   "Goto"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton BN_GSB 
      Caption         =   "GSB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      ToolTipText     =   "Go to subroutine"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton BN_DIV 
      Caption         =   "¸"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton BN_9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton BN_8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton BN_7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton BN_F 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton BN_E 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton BN_D 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton BN_C 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton BN_B 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton BN_A 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Display 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0.0000"
      Top             =   240
      Width           =   4680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "SQTx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   210
      Index           =   14
      Left            =   2760
      TabIndex        =   54
      ToolTipText     =   "Square root"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "ISZ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   13
      Left            =   2160
      TabIndex        =   53
      ToolTipText     =   "Increment skip on zero"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "DSZ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   12
      Left            =   1560
      TabIndex        =   52
      ToolTipText     =   "Decrement skip on zero"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "----------  clear  -----------"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   48
      Left            =   1680
      TabIndex        =   109
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "F?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   18
      Left            =   5160
      TabIndex        =   58
      ToolTipText     =   "Flag value"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "CF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   17
      Left            =   4560
      TabIndex        =   57
      ToolTipText     =   "Clear flag"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "SF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   16
      Left            =   3960
      TabIndex        =   56
      ToolTipText     =   "Set flag"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "---------  set compl  --------"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   50
      Left            =   3960
      TabIndex        =   110
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "2's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   63
      Left            =   4560
      TabIndex        =   100
      ToolTipText     =   "Set 2's complement"
      Top             =   2370
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "RRCn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   5
      Left            =   3360
      TabIndex        =   45
      ToolTipText     =   "Rotate right with carry n times"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "RLCn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   4
      Left            =   2760
      TabIndex        =   44
      ToolTipText     =   "Rotate left with carry n times"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "RRC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   3
      Left            =   2160
      TabIndex        =   43
      ToolTipText     =   "Rotate right with carry"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "RLC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   2
      Left            =   1560
      TabIndex        =   42
      ToolTipText     =   "Rotate left with carry"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   5520
      Picture         =   "calc.frx":030A
      Top             =   120
      Width           =   555
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   120
      Top             =   0
      Width           =   6375
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   3735
      Index           =   1
      Left            =   6480
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   3735
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "E M M E T -  G R A Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   480
      TabIndex        =   111
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   2535
      Left            =   240
      Top             =   1080
      Width           =   6135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   71
      Left            =   5760
      TabIndex        =   108
      ToolTipText     =   "Logical OR"
      Top             =   2960
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "EEX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   70
      Left            =   5160
      TabIndex        =   107
      ToolTipText     =   "Exponent"
      Top             =   2960
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   69
      Left            =   4440
      TabIndex        =   106
      ToolTipText     =   "Displays status"
      Top             =   2955
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "MEM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   68
      Left            =   3960
      TabIndex        =   105
      ToolTipText     =   "Available memory"
      Top             =   2960
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "FLOAT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   67
      Left            =   2760
      TabIndex        =   104
      ToolTipText     =   "Floating point mode"
      Top             =   2960
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "WSIZE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   66
      Left            =   2160
      TabIndex        =   103
      ToolTipText     =   "Word size"
      Top             =   2960
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "NOT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   65
      Left            =   5760
      TabIndex        =   102
      ToolTipText     =   "Logical NOT"
      Top             =   2360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "unsgn"
      DataSource      =   "unsgn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   145
      Index           =   64
      Left            =   5160
      TabIndex        =   101
      ToolTipText     =   "Set unsigned"
      Top             =   2350
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "1's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   62
      Left            =   3960
      TabIndex        =   99
      ToolTipText     =   "Set 1's complement"
      Top             =   2370
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "window"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   61
      Left            =   3360
      TabIndex        =   98
      ToolTipText     =   "Window size"
      Top             =   2360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "prefix"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   145
      Index           =   60
      Left            =   2760
      TabIndex        =   97
      ToolTipText     =   "Clear prefix"
      Top             =   2350
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "reg"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   59
      Left            =   2160
      TabIndex        =   96
      ToolTipText     =   "Clear register"
      Top             =   2370
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "pgrm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   145
      Index           =   58
      Left            =   1560
      TabIndex        =   95
      ToolTipText     =   "Clear program"
      Top             =   2350
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   57
      Left            =   960
      TabIndex        =   94
      ToolTipText     =   "Index"
      Top             =   2360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "(i)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   56
      Left            =   360
      TabIndex        =   93
      ToolTipText     =   "Relative index"
      Top             =   2355
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "AND"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   55
      Left            =   5760
      TabIndex        =   92
      ToolTipText     =   "Logical AND"
      Top             =   1760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "B?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   54
      Left            =   5160
      TabIndex        =   91
      ToolTipText     =   "Bit value"
      Top             =   1760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "CB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   53
      Left            =   4560
      TabIndex        =   90
      ToolTipText     =   "Clear bit"
      Top             =   1760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "SB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   52
      Left            =   3960
      TabIndex        =   89
      ToolTipText     =   "Set Bit"
      Top             =   1760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "------------------  show ---------------------"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   49
      Left            =   1560
      TabIndex        =   88
      Top             =   1705
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "x:i"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   47
      Left            =   960
      TabIndex        =   87
      ToolTipText     =   "Exchange X I"
      Top             =   1760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "x:(i)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   46
      Left            =   360
      TabIndex        =   86
      ToolTipText     =   "Exchange X index relative"
      Top             =   1760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "XOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   45
      Left            =   5760
      TabIndex        =   85
      ToolTipText     =   "Exclusive OR"
      Top             =   1160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "RMD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   44
      Left            =   5160
      TabIndex        =   84
      ToolTipText     =   "Remainder"
      Top             =   1160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "MASKR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   43
      Left            =   4560
      TabIndex        =   83
      ToolTipText     =   "Mask right"
      Top             =   1160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "MASKL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   42
      Left            =   3960
      TabIndex        =   82
      ToolTipText     =   "Mask left"
      Top             =   1160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "RRn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   41
      Left            =   3360
      TabIndex        =   81
      ToolTipText     =   "Rotate right n times"
      Top             =   1160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "RLn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   40
      Left            =   2760
      TabIndex        =   80
      ToolTipText     =   "Rotate left n times"
      Top             =   1160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "RR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   39
      Left            =   2160
      TabIndex        =   79
      ToolTipText     =   "Rotate right"
      Top             =   1160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "RL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   38
      Left            =   1560
      TabIndex        =   78
      ToolTipText     =   "Rotate Left"
      Top             =   1160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "SR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   37
      Left            =   960
      TabIndex        =   77
      ToolTipText     =   "Shift right"
      Top             =   1160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "SL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Index           =   36
      Left            =   360
      TabIndex        =   76
      ToolTipText     =   "Shift left"
      Top             =   1160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "x=0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   35
      Left            =   5760
      TabIndex        =   75
      ToolTipText     =   "if x = 0"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "x=y"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   34
      Left            =   5160
      TabIndex        =   74
      ToolTipText     =   "if x = y"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "x!=0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   33
      Left            =   4560
      TabIndex        =   73
      ToolTipText     =   "if x != 0"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "x!=y"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   32
      Left            =   3960
      TabIndex        =   72
      ToolTipText     =   "if x != y"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "LSTx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   31
      Left            =   3360
      TabIndex        =   71
      ToolTipText     =   "Last X"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   30
      Left            =   2760
      TabIndex        =   70
      ToolTipText     =   "Shift display right"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   29
      Left            =   2160
      TabIndex        =   69
      ToolTipText     =   "Shift display left"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "x>0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   28
      Left            =   5760
      TabIndex        =   68
      ToolTipText     =   "if x > 0"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "x>y"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   27
      Left            =   5160
      TabIndex        =   67
      ToolTipText     =   "if x > y"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "x<0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   26
      Left            =   4560
      TabIndex        =   66
      ToolTipText     =   "if x < 0"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "x<=y"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   25
      Left            =   3960
      TabIndex        =   65
      ToolTipText     =   "if x <= y"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "CLx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   24
      Left            =   2760
      TabIndex        =   64
      ToolTipText     =   "Clear x"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "PSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   23
      Left            =   2160
      TabIndex        =   63
      ToolTipText     =   "Pause"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "R up"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   22
      Left            =   1560
      TabIndex        =   62
      ToolTipText     =   "Roll up stack"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "BST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   21
      Left            =   960
      TabIndex        =   61
      ToolTipText     =   "Back step"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "P/R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   20
      Left            =   360
      TabIndex        =   60
      ToolTipText     =   "Pause/Run"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "DBLx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   19
      Left            =   5760
      TabIndex        =   59
      ToolTipText     =   "Double multiply"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   15
      Left            =   3360
      TabIndex        =   55
      ToolTipText     =   "Inverse"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "LBL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   11
      Left            =   960
      TabIndex        =   51
      ToolTipText     =   "Label"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "RTN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   10
      Left            =   360
      TabIndex        =   50
      ToolTipText     =   "Return"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "DBL/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   9
      Left            =   5760
      TabIndex        =   49
      ToolTipText     =   "Double divide"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "DBLR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   8
      Left            =   5160
      TabIndex        =   48
      ToolTipText     =   "Double Rotate???"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "ABS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   7
      Left            =   4560
      TabIndex        =   47
      ToolTipText     =   "Absolute value"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "#B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   6
      Left            =   3960
      TabIndex        =   46
      ToolTipText     =   "Number of bits"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "ASR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   1
      Left            =   960
      TabIndex        =   41
      ToolTipText     =   "Arithmetic shift right"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "LJ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   135
      Index           =   0
      Left            =   360
      TabIndex        =   40
      ToolTipText     =   "Left justify"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   2775
      Left            =   120
      Top             =   960
      Width           =   6375
   End
   Begin VB.Menu CM_FILE 
      Caption         =   "File"
      Begin VB.Menu CM_FILE_NEW 
         Caption         =   "New"
         Enabled         =   0   'False
      End
      Begin VB.Menu CM_FILE_OPEN 
         Caption         =   "Open"
         Enabled         =   0   'False
      End
      Begin VB.Menu CM_FILE_SAVE 
         Caption         =   "Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu CM_FILE_SAVEAS 
         Caption         =   "Save as"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu CM_FILE_PRINT 
         Caption         =   "Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu CM_FILE_PAGE 
         Caption         =   "Page setup"
         Enabled         =   0   'False
      End
      Begin VB.Menu CM_FILE_PRINTER 
         Caption         =   "Printer setup"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu CM_FILE_EXIT 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu CM_EDIT 
      Caption         =   "Edit"
      Begin VB.Menu CM_EDIT_COPY 
         Caption         =   "Copy"
      End
      Begin VB.Menu CM_EDIT_PASTE 
         Caption         =   "Paste"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu CM_VIEW 
      Caption         =   "View"
      Begin VB.Menu CM_VIEW_FLOAT 
         Caption         =   "Floating point mode"
         Checked         =   -1  'True
      End
      Begin VB.Menu CM_VIEW_DEC 
         Caption         =   "Decimal mode"
      End
      Begin VB.Menu CM_VIEW_HEX 
         Caption         =   "Hexidecimal mode"
      End
      Begin VB.Menu CM_VIEW_OCT 
         Caption         =   "Octal mode"
      End
      Begin VB.Menu CM_VIEW_BIN 
         Caption         =   "Binary mode"
      End
   End
   Begin VB.Menu CM_OPTIONS 
      Caption         =   "Options"
      Begin VB.Menu CM_OPTION_8 
         Caption         =   "8 bit word"
      End
      Begin VB.Menu CM_OPTION_16 
         Caption         =   "16 bit word"
      End
      Begin VB.Menu CM_OPTION_32 
         Caption         =   "32 bit word"
         Checked         =   -1  'True
      End
      Begin VB.Menu CM_OPTION_64 
         Caption         =   "64 bit word"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu CM_OPTION_1S 
         Caption         =   "1's complement"
      End
      Begin VB.Menu CM_OPTION_2S 
         Caption         =   "2's complement"
         Checked         =   -1  'True
      End
      Begin VB.Menu CM_OPTION_UNS 
         Caption         =   "Unsigned"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu CM_OPTION_FLAG0 
         Caption         =   "User Flag 0"
      End
      Begin VB.Menu CM_OPTION_FLAG1 
         Caption         =   "User Flag 1"
      End
      Begin VB.Menu CM_OPTION_FLAG2 
         Caption         =   "User Flag 2"
      End
      Begin VB.Menu CM_OPTION_FLAG3 
         Caption         =   "Leading Zeros"
      End
      Begin VB.Menu CM_OPTION_FLAG4 
         Caption         =   "Carry Flag"
      End
      Begin VB.Menu CM_OPTION_FLAG5 
         Caption         =   "Overlow"
      End
   End
   Begin VB.Menu CM_HELP 
      Caption         =   "Help"
      Begin VB.Menu CM_HELP_CONTENTS 
         Caption         =   "Contents"
      End
      Begin VB.Menu CM_HELP_ABOUT 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim func_key As Integer
Private Sub Form_load()
    Dim i As Integer
    Dim zero As BigInt
    
    ' Process key events at the "form level" before passing on to other
    ' window objects
    KeyPreview = True
    
    ' Set the starting conditions
    mode = FLOAT_MODE
    word_size = 32
    dpoints = 4
    comp = 2

    ' build the masks
    word_mask = BigWordMask(word_size)
    sign_mask = BigSignMask(word_size)
    
    ' check the clipboard
    If Clipboard.GetFormat(vbCFText) = True Then
        CM_EDIT_PASTE.Enabled = True
    Else
        CM_EDIT_PASTE.Enabled = False
    End If
End Sub
Private Sub Form_KeyUp(Keycode As Integer, Shift As Integer)
    ' Detect the release of the CTRL and ALT calc_key
    
    If Keycode = vbKeyControl Then
        func_key = 0
        DispFlags DFLAG_FKEY, False
    End If
    If Keycode = vbKeyMenu Then
        func_key = 0
        DispFlags DFLAG_GKEY, False
    End If
End Sub
Private Sub Form_KeyDown(Keycode As Integer, Shift As Integer)
    ' Detect the press of the CTRL and ALT calc_key.  Also handles
    ' normal keyboard entry from the numeric keypad
    
    Select Case Keycode
        Case vbKeyControl
            DispFlags DFLAG_FKEY, True
            func_key = CTRL
        Case vbKeyMenu
            DispFlags DFLAG_GKEY, True
            func_key = ALT
            ' We must zap the keycode value for ALT to prevent it from
            ' activating the command menu
            Keycode = 0
        Case vbKey0 To vbKey9
            Call calc_key(Keycode)
        Case vbKeyNumpad0 To vbKeyNumpad9
            Call calc_key(Keycode - 48)
        Case vbKeyA To vbKeyF
            Call calc_key(Keycode + 32)
        Case vbKeyAdd
            Call calc_key(K_PLUS)
        Case vbKeyMultiply
            Call calc_key(K_MULT)
        Case vbKeySeparator, K_ENTER
            Call calc_key(K_ENTER)
        Case vbKeySubtract
            Call calc_key(K_MINUS)
        Case vbKeyDecimal, K_DP
            Call calc_key(K_DP)
        Case vbKeyDivide
            Call calc_key(K_DIV)
        ' allow these keys to pass
        Case vbKeyShift
        Case vbKeyNumlock
        ' beep on everything else
        Case Else
             Beep
    End Select
End Sub
'
' all of the calc_key
'
Private Sub BN_A_Click()
    calc_key (K_A + func_key)
End Sub
Private Sub BN_B_Click()
    calc_key (K_B + func_key)
End Sub
Private Sub BN_C_Click()
    calc_key (K_C + func_key)
End Sub
Private Sub BN_D_Click()
    calc_key (K_D + func_key)
End Sub
Private Sub BN_E_Click()
    calc_key (K_E + func_key)
End Sub
Private Sub BN_F_Click()
    calc_key (K_F + func_key)
End Sub
Private Sub BN_7_Click()
    calc_key (K_7 + func_key)
End Sub
Private Sub BN_8_Click()
    calc_key (K_8 + func_key)
End Sub
Private Sub BN_9_Click()
    calc_key (K_9 + func_key)
End Sub
Private Sub BN_DIV_Click()
    calc_key (K_DIV + func_key)
End Sub
Private Sub BN_GSB_Click()
    calc_key (K_GSB + func_key)
End Sub
Private Sub BN_GTO_Click()
    calc_key (K_GTO + func_key)
End Sub
Private Sub BN_HEX_Click()
    calc_key (K_HEX + func_key)
End Sub
Private Sub BN_DEC_Click()
   calc_key (K_DEC + func_key)
End Sub
Private Sub BN_OCT_Click()
    calc_key (K_OCT + func_key)
End Sub
Private Sub BN_BIN_Click()
    calc_key (K_BIN + func_key)
End Sub
Private Sub BN_4_Click()
    calc_key (K_4 + func_key)
End Sub
Private Sub BN_5_Click()
    calc_key (K_5 + func_key)
End Sub
Private Sub BN_6_Click()
    calc_key (K_6 + func_key)
End Sub
Private Sub BN_MULT_Click()
    calc_key (K_MULT + func_key)
End Sub
Private Sub BN_RS_Click()
    calc_key (K_RS + func_key)
End Sub
Private Sub BN_SST_Click()
    calc_key (K_SST + func_key)
End Sub
Private Sub BN_ROL_Click()
    calc_key (K_ROL + func_key)
End Sub
Private Sub BN_XY_Click()
    calc_key (K_XY + func_key)
End Sub
Private Sub BN_BSP_Click()
    calc_key (K_BSP + func_key)
End Sub
Private Sub BN_ENTER_Click()
    calc_key (K_ENTER + func_key)
End Sub
Private Sub BN_1_Click()
    calc_key (K_1 + func_key)
End Sub
Private Sub BN_2_Click()
    calc_key (K_2 + func_key)
End Sub
Private Sub BN_3_Click()
    calc_key (K_3 + func_key)
End Sub
Private Sub BN_MINUS_Click()
    calc_key (K_MINUS + func_key)
End Sub
Private Sub BN_ON_Click()
    Unload Me
End Sub
Private Sub BN_FKEY_Click()
    DispFlags DFLAG_FKEY, True
    calc_key (K_FKEY)
End Sub
Private Sub BN_GKEY_Click()
    DispFlags DFLAG_GKEY, True
    calc_key (K_GKEY)
End Sub
Private Sub BN_STO_Click()
    calc_key (K_STO + func_key)
End Sub
Private Sub BN_RCL_Click()
    calc_key (K_RCL + func_key)
End Sub
Private Sub BN_0_Click()
    calc_key (K_0 + func_key)
End Sub
Private Sub BN_DP_Click()
    calc_key (K_DP + func_key)
End Sub
Private Sub BN_CHS_Click()
    calc_key (K_CHS + func_key)
End Sub
Private Sub BN_PLUS_Click()
    calc_key (K_PLUS + func_key)
End Sub
'
' the command menu
'
Private Sub CM_FILE_EXIT_Click()
    Unload Me
End Sub
Private Sub CM_EDIT_COPY_Click()
    If mode = FLOAT_MODE Then
        Clipboard.SetText (RTrim(Display.Text))
    Else
        Clipboard.SetText (LTrim(Display.Text))
    End If
End Sub
Private Sub CM_EDIT_PASTE_Click()
    ' the problem here is that we no way to determine if the text on
    ' the clipboard is in the proper format
    Dim i As Integer
    Dim buf As String
    Dim c As String
    Dim odd As Integer
    Dim f As Double
    Dim x As BigInt
    
    buf = Trim(Clipboard.GetText())
    If Len(buf) = 0 Then
        Exit Sub
    End If
    odd = 0
    For i = 1 To Len(buf)
        c = Mid(buf, i, 1)
        Select Case mode
            Case FLOAT_MODE
                If InStr("-0123456789.eE", c) = 0 Then
                    odd = 1
                    Exit For
                End If
            Case DEC_MODE
                If InStr("-0123456789", c) = 0 Then
                    odd = 1
                    Exit For
                End If
            Case HEX_MODE
                If InStr("0123456789abcefABCDEF", c) = 0 Then
                    odd = 1
                    Exit For
                End If
             Case OCT_MODE
                If InStr("01234567", c) = 0 Then
                    odd = 1
                    Exit For
                End If
             Case BIN_MODE
                If InStr(" 01", c) = 0 Then
                    odd = 1
                    Exit For
                End If
        End Select
    Next
    ' all we do is beep if something is fishy
    If odd = 1 Then
        Beep
    End If
    
    If mode = FLOAT_MODE Then
        f = Convert_Input(buf)
        push (f)
    Else
        x = Convert_Input64(buf, mode)
        push64 x
    End If
    calc_key (0)
End Sub
Public Sub CM_VIEW_FLOAT_Click()
    CM_VIEW_FLOAT.Checked = True
    CM_VIEW_DEC.Checked = False
    CM_VIEW_HEX.Checked = False
    CM_VIEW_OCT.Checked = False
    CM_VIEW_BIN.Checked = False
    calc_key (0)
    mode = FLOAT_MODE
    calc_key (0)
End Sub
Public Sub CM_VIEW_DEC_Click()
    mode = DEC_MODE
    CM_VIEW_FLOAT.Checked = False
    CM_VIEW_DEC.Checked = True
    CM_VIEW_HEX.Checked = False
    CM_VIEW_OCT.Checked = False
    CM_VIEW_BIN.Checked = False
    calc_key (0)
    mode = DEC_MODE
    calc_key (0)
End Sub
Public Sub CM_VIEW_HEX_Click()
    mode = HEX_MODE
    CM_VIEW_FLOAT.Checked = False
    CM_VIEW_DEC.Checked = False
    CM_VIEW_HEX.Checked = True
    CM_VIEW_OCT.Checked = False
    CM_VIEW_BIN.Checked = False
    calc_key (0)
    mode = HEX_MODE
    calc_key (0)
End Sub
Public Sub CM_VIEW_OCT_Click()
    mode = OCT_MODE
    CM_VIEW_FLOAT.Checked = False
    CM_VIEW_DEC.Checked = False
    CM_VIEW_HEX.Checked = False
    CM_VIEW_OCT.Checked = True
    CM_VIEW_BIN.Checked = False
    calc_key (0)
    mode = OCT_MODE
    calc_key (0)
End Sub
Public Sub CM_VIEW_BIN_Click()
    mode = BIN_MODE
    CM_VIEW_FLOAT.Checked = False
    CM_VIEW_DEC.Checked = False
    CM_VIEW_HEX.Checked = False
    CM_VIEW_OCT.Checked = False
    CM_VIEW_BIN.Checked = True
    calc_key (0)
    mode = BIN_MODE
    calc_key (0)
End Sub
Private Sub CM_OPTION_8_Click()
    CM_OPTION_8.Checked = True
    CM_OPTION_16.Checked = False
    CM_OPTION_32.Checked = False
    CM_OPTION_64.Checked = False
    calc_key (0)
    word_mask = BigWordMask(8)
    sign_mask = BigSignMask(8)
    If word_size > 8 Then
        MaskAll
    End If
    word_size = 8
    calc_key (0)
End Sub
Private Sub CM_OPTION_16_Click()
    CM_OPTION_8.Checked = False
    CM_OPTION_16.Checked = True
    CM_OPTION_32.Checked = False
    CM_OPTION_64.Checked = False
    calc_key (0)
    word_mask = BigWordMask(16)
    sign_mask = BigSignMask(16)
    If word_size > 16 Then
        MaskAll
    End If
    word_size = 16
    calc_key (0)
End Sub
Private Sub CM_OPTION_32_Click()
    CM_OPTION_8.Checked = False
    CM_OPTION_16.Checked = False
    CM_OPTION_32.Checked = True
    CM_OPTION_64.Checked = False
    calc_key (0)
    word_mask = BigWordMask(32)
    sign_mask = BigSignMask(32)
    If word_size > 32 Then
        MaskAll
    End If
    word_size = 32
    calc_key (0)
End Sub
Private Sub CM_OPTION_64_Click()
    CM_OPTION_8.Checked = False
    CM_OPTION_16.Checked = False
    CM_OPTION_32.Checked = False
    CM_OPTION_64.Checked = True
    calc_key (0)
    word_mask = BigWordMask(64)
    sign_mask = BigSignMask(64)
    word_size = 64
    calc_key (0)
End Sub
Private Sub CM_OPTION_1S_Click()
    CM_OPTION_1S.Checked = True
    CM_OPTION_2S.Checked = False
    CM_OPTION_UNS.Checked = False
    calc_key (0)
    comp = 1
    calc_key (0)
End Sub
Private Sub CM_OPTION_2S_Click()
    CM_OPTION_1S.Checked = False
    CM_OPTION_2S.Checked = True
    CM_OPTION_UNS.Checked = False
    calc_key (0)
    comp = 2
    calc_key (0)
End Sub
Private Sub CM_OPTION_UNS_Click()
    CM_OPTION_1S.Checked = False
    CM_OPTION_2S.Checked = False
    CM_OPTION_UNS.Checked = True
    calc_key (0)
    comp = 0
    calc_key (0)
End Sub
Private Sub CM_HELP_CONTENTS_Click()
    With CommonDialog1
        ' .HelpCommand = cdlHelpContents
        ' .HelpFile = "wrpn.hlp"
        ' .ShowHelp
    End With
End Sub
Private Sub CM_HELP_ABOUT_Click()
    About.Show
End Sub
Private Sub CM_OPTION_FLAG0_Click()
    If flag(0) = True Then
        CM_OPTION_FLAG0.Checked = False
        flag(0) = False
    Else
        CM_OPTION_FLAG0.Checked = True
        flag(0) = True
    End If
End Sub
Private Sub CM_OPTION_FLAG1_Click()
    If flag(1) = True Then
        CM_OPTION_FLAG1.Checked = False
        flag(1) = False
    Else
        CM_OPTION_FLAG1.Checked = True
        flag(1) = True
    End If
End Sub
Private Sub CM_OPTION_FLAG2_Click()
    If flag(2) = True Then
        CM_OPTION_FLAG2.Checked = False
        flag(2) = False
    Else
        CM_OPTION_FLAG2.Checked = True
        flag(2) = True
    End If
End Sub
Private Sub CM_OPTION_FLAG3_Click()
    If flag(3) = True Then
        CM_OPTION_FLAG3.Checked = False
        flag(3) = False
    Else
        CM_OPTION_FLAG3.Checked = True
        flag(3) = True
    End If
    calc_key (0)
End Sub
Private Sub CM_OPTION_FLAG4_Click()
    If flag(4) = True Then
        CM_OPTION_FLAG4.Checked = False
        DispFlags DFLAG_CARRYBIT, False
    Else
        CM_OPTION_FLAG4.Checked = True
        DispFlags DFLAG_CARRYBIT, True
    End If
End Sub
Private Sub CM_OPTION_FLAG5_Click()
    If flag(5) = True Then
        CM_OPTION_FLAG5.Checked = False
        DispFlags DFLAG_OVERFLOW, False
    Else
        CM_OPTION_FLAG5.Checked = True
        DispFlags DFLAG_OVERFLOW, True
    End If
End Sub

