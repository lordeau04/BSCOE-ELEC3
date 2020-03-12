VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command47 
      Caption         =   "AC"
      Height          =   375
      Left            =   3960
      TabIndex        =   47
      Top             =   4920
      Width           =   750
   End
   Begin VB.CommandButton Command46 
      Caption         =   "DEL"
      Height          =   375
      Left            =   3000
      TabIndex        =   46
      Top             =   4920
      Width           =   750
   End
   Begin VB.CommandButton Command45 
      Caption         =   "9"
      Height          =   375
      Left            =   2040
      TabIndex        =   45
      Top             =   4920
      Width           =   750
   End
   Begin VB.CommandButton Command44 
      Caption         =   "8"
      Height          =   375
      Left            =   1080
      TabIndex        =   44
      Top             =   4920
      Width           =   750
   End
   Begin VB.CommandButton Command43 
      Caption         =   "7"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   4920
      Width           =   750
   End
   Begin VB.CommandButton Command42 
      Caption         =   "/"
      Height          =   375
      Left            =   3960
      TabIndex        =   42
      Top             =   5520
      Width           =   750
   End
   Begin VB.CommandButton Command41 
      Caption         =   "*"
      Height          =   375
      Left            =   3000
      TabIndex        =   41
      Top             =   5520
      Width           =   750
   End
   Begin VB.CommandButton Command40 
      Caption         =   "6"
      Height          =   375
      Left            =   2040
      TabIndex        =   40
      Top             =   5520
      Width           =   750
   End
   Begin VB.CommandButton Command39 
      Caption         =   "5"
      Height          =   375
      Left            =   1080
      TabIndex        =   39
      Top             =   5520
      Width           =   750
   End
   Begin VB.CommandButton Command38 
      Caption         =   "4"
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   5520
      Width           =   750
   End
   Begin VB.CommandButton Command37 
      Caption         =   "-"
      Height          =   375
      Left            =   3960
      TabIndex        =   37
      Top             =   6120
      Width           =   750
   End
   Begin VB.CommandButton Command36 
      Caption         =   "+"
      Height          =   375
      Left            =   3000
      TabIndex        =   36
      Top             =   6120
      Width           =   750
   End
   Begin VB.CommandButton Command35 
      Caption         =   "3"
      Height          =   375
      Left            =   2040
      TabIndex        =   35
      Top             =   6120
      Width           =   750
   End
   Begin VB.CommandButton Command34 
      Caption         =   "2"
      Height          =   375
      Left            =   1080
      TabIndex        =   34
      Top             =   6120
      Width           =   750
   End
   Begin VB.CommandButton Command33 
      Caption         =   "1"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   6120
      Width           =   750
   End
   Begin VB.CommandButton Command32 
      Caption         =   "="
      Height          =   375
      Left            =   3960
      TabIndex        =   32
      Top             =   6720
      Width           =   750
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Ans"
      Height          =   375
      Left            =   3000
      TabIndex        =   31
      Top             =   6720
      Width           =   750
   End
   Begin VB.CommandButton Command30 
      Caption         =   "x10^x"
      Height          =   375
      Left            =   2040
      TabIndex        =   30
      Top             =   6720
      Width           =   750
   End
   Begin VB.CommandButton Command29 
      Caption         =   "."
      Height          =   375
      Left            =   1080
      TabIndex        =   29
      Top             =   6720
      Width           =   750
   End
   Begin VB.CommandButton Command28 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   6720
      Width           =   750
   End
   Begin VB.CommandButton Command27 
      Caption         =   "OFF"
      Height          =   375
      Left            =   3960
      TabIndex        =   27
      Top             =   7440
      Width           =   750
   End
   Begin VB.CommandButton Command26 
      Caption         =   "ln"
      Height          =   375
      Left            =   3960
      TabIndex        =   26
      Top             =   3240
      Width           =   550
   End
   Begin VB.CommandButton Command25 
      Caption         =   "log"
      Height          =   375
      Left            =   3240
      TabIndex        =   25
      Top             =   3240
      Width           =   550
   End
   Begin VB.CommandButton Command24 
      Caption         =   "tan"
      Height          =   375
      Left            =   3960
      TabIndex        =   24
      Top             =   3720
      Width           =   550
   End
   Begin VB.CommandButton Command23 
      Caption         =   "cos"
      Height          =   375
      Left            =   3240
      TabIndex        =   23
      Top             =   3720
      Width           =   550
   End
   Begin VB.CommandButton Command22 
      Caption         =   "M+"
      Height          =   375
      Left            =   3960
      TabIndex        =   22
      Top             =   4200
      Width           =   550
   End
   Begin VB.CommandButton Command21 
      Caption         =   "S<>D"
      Height          =   375
      Left            =   3240
      TabIndex        =   21
      Top             =   4200
      Width           =   550
   End
   Begin VB.CommandButton Command20 
      Caption         =   "x^[]"
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   3240
      Width           =   550
   End
   Begin VB.CommandButton Command19 
      Caption         =   "x^2"
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   3240
      Width           =   550
   End
   Begin VB.CommandButton Command18 
      Caption         =   "sin"
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   3720
      Width           =   550
   End
   Begin VB.CommandButton Command17 
      Caption         =   "hyp"
      Height          =   375
      Left            =   1800
      TabIndex        =   17
      Top             =   3720
      Width           =   550
   End
   Begin VB.CommandButton Command16 
      Caption         =   ")"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   4200
      Width           =   550
   End
   Begin VB.CommandButton Command15 
      Caption         =   "("
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   4200
      Width           =   550
   End
   Begin VB.CommandButton Command14 
      Caption         =   "integ"
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   2760
      Width           =   550
   End
   Begin VB.CommandButton Command13 
      Caption         =   "sq root"
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   3240
      Width           =   550
   End
   Begin VB.CommandButton Command12 
      Caption         =   ".' """
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   3720
      Width           =   550
   End
   Begin VB.CommandButton Command11 
      Caption         =   "ENG"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   4200
      Width           =   550
   End
   Begin VB.CommandButton Command10 
      Caption         =   "CALC"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   550
   End
   Begin VB.CommandButton Command9 
      Caption         =   "x/x"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   550
   End
   Begin VB.CommandButton Command8 
      Caption         =   "( - )"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   550
   End
   Begin VB.CommandButton Command7 
      Caption         =   "RCL"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   550
   End
   Begin VB.CommandButton Command6 
      Caption         =   "x^-1"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   2760
      Width           =   550
   End
   Begin VB.CommandButton Command5 
      Caption         =   "log[]"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   2760
      Width           =   550
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SHIFT"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ALPHA"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   2040
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MODE"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2040
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ON"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label Label3 
      Caption         =   "BS CpE 5-4"
      Height          =   255
      Left            =   360
      TabIndex        =   49
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Submitted by: Danreb S. Guiyab"
      Height          =   255
      Left            =   360
      TabIndex        =   48
      Top             =   7440
      Width           =   2655
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command28_Click()
Label1.Caption = "0"
End Sub

Private Sub Command33_Click()
Label1.Caption = "1"
End Sub

Private Sub Command34_Click()
Label1.Caption = "2"
End Sub

Private Sub Command35_Click()
Label1.Caption = "3"
End Sub

Private Sub Command38_Click()
Label1.Caption = "4"
End Sub

Private Sub Command39_Click()
Label1.Caption = "5"
End Sub

Private Sub Command40_Click()
Label1.Caption = "6"
End Sub

Private Sub Command43_Click()
Label1.Caption = "7"
End Sub

Private Sub Command44_Click()
Label1.Caption = "8"
End Sub

Private Sub Command45_Click()
Label1.Caption = "9"
End Sub

Private Sub Command46_Click()
Label1.Caption = ""
End Sub
