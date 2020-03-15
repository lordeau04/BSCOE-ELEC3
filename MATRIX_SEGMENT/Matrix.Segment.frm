VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5500
      Top             =   7560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "STOP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   26
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   25
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Label Label27 
      Caption         =   "Submitted by: Danreb S. Guiyab CpE 5-4"
      Height          =   255
      Left            =   1800
      TabIndex        =   28
      Top             =   7680
      Width           =   3015
   End
   Begin VB.Label Label26 
      Caption         =   "Label26"
      Height          =   375
      Left            =   600
      TabIndex        =   27
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label Label25 
      Height          =   975
      Left            =   4200
      TabIndex        =   24
      Top             =   5640
      Width           =   705
   End
   Begin VB.Label Label24 
      Height          =   975
      Left            =   3240
      TabIndex        =   23
      Top             =   5640
      Width           =   705
   End
   Begin VB.Label Label23 
      Height          =   975
      Left            =   2280
      TabIndex        =   22
      Top             =   5640
      Width           =   705
   End
   Begin VB.Label Label22 
      Height          =   975
      Left            =   1320
      TabIndex        =   21
      Top             =   5640
      Width           =   705
   End
   Begin VB.Label Label21 
      Height          =   975
      Left            =   360
      TabIndex        =   20
      Top             =   5640
      Width           =   705
   End
   Begin VB.Label Label20 
      Height          =   975
      Left            =   4200
      TabIndex        =   19
      Top             =   4320
      Width           =   705
   End
   Begin VB.Label Label19 
      Height          =   975
      Left            =   3240
      TabIndex        =   18
      Top             =   4320
      Width           =   705
   End
   Begin VB.Label Label18 
      Height          =   975
      Left            =   2280
      TabIndex        =   17
      Top             =   4320
      Width           =   705
   End
   Begin VB.Label Label17 
      Height          =   975
      Left            =   1320
      TabIndex        =   16
      Top             =   4320
      Width           =   705
   End
   Begin VB.Label Label16 
      Height          =   975
      Left            =   360
      TabIndex        =   15
      Top             =   4320
      Width           =   705
   End
   Begin VB.Label Label15 
      Height          =   975
      Left            =   4200
      TabIndex        =   14
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label Label14 
      Height          =   975
      Left            =   3240
      TabIndex        =   13
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label Label13 
      Height          =   975
      Left            =   2280
      TabIndex        =   12
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label Label12 
      Height          =   975
      Left            =   1320
      TabIndex        =   11
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label Label11 
      Height          =   975
      Left            =   360
      TabIndex        =   10
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label Label10 
      Height          =   975
      Left            =   4200
      TabIndex        =   9
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label Label9 
      Height          =   975
      Left            =   3240
      TabIndex        =   8
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label Label8 
      Height          =   975
      Left            =   2280
      TabIndex        =   7
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label Label7 
      Height          =   975
      Left            =   1320
      TabIndex        =   6
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label Label6 
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label Label5 
      Height          =   975
      Left            =   4200
      TabIndex        =   4
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label4 
      Height          =   975
      Left            =   3240
      TabIndex        =   3
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label3 
      Height          =   975
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label2 
      Height          =   975
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label1 
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   700
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Long

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub
Private Sub Command2_Click()
Timer1.Enabled = False
End Sub
Private Sub Form_Load()
Timer1.Enabled = False
c = -1
End Sub

Private Sub Timer1_Timer()
c = c + 1
Label26.Caption = c
If c = 0 Then
Label1.BackColor = vbBlack
Label2.BackColor = vbBlack
Label3.BackColor = vbBlack
Label4.BackColor = vbBlack
Label5.BackColor = vbBlack
Label6.BackColor = vbBlack
Label7.BackColor = &H8000000F
Label8.BackColor = &H8000000F
Label9.BackColor = &H8000000F
Label10.BackColor = vbBlack
Label11.BackColor = vbBlack
Label12.BackColor = &H8000000F
Label13.BackColor = &H8000000F
Label14.BackColor = &H8000000F
Label15.BackColor = vbBlack
Label16.BackColor = vbBlack
Label17.BackColor = &H8000000F
Label18.BackColor = &H8000000F
Label19.BackColor = &H8000000F
Label20.BackColor = vbBlack
Label21.BackColor = vbBlack
Label22.BackColor = vbBlack
Label23.BackColor = vbBlack
Label24.BackColor = vbBlack
Label25.BackColor = vbBlack
End If

If c = 1 Then
Label1.BackColor = &H8000000F
Label2.BackColor = vbBlack
Label3.BackColor = vbBlack
Label4.BackColor = &H8000000F
Label5.BackColor = &H8000000F
Label6.BackColor = vbBlack
Label7.BackColor = &H8000000F
Label8.BackColor = vbBlack
Label9.BackColor = &H8000000F
Label10.BackColor = &H8000000F
Label11.BackColor = &H8000000F
Label12.BackColor = &H8000000F
Label13.BackColor = vbBlack
Label14.BackColor = &H8000000F
Label15.BackColor = &H8000000F
Label16.BackColor = &H8000000F
Label17.BackColor = &H8000000F
Label18.BackColor = vbBlack
Label19.BackColor = &H8000000F
Label20.BackColor = &H8000000F
Label21.BackColor = &H8000000F
Label22.BackColor = &H8000000F
Label23.BackColor = vbBlack
Label24.BackColor = &H8000000F
Label25.BackColor = &H8000000F
End If

If c = 2 Then
Label1.BackColor = &H8000000F
Label2.BackColor = vbBlack
Label3.BackColor = vbBlack
Label4.BackColor = vbBlack
Label5.BackColor = &H8000000F
Label6.BackColor = vbBlack
Label7.BackColor = &H8000000F
Label8.BackColor = &H8000000F
Label9.BackColor = &H8000000F
Label10.BackColor = vbBlack
Label11.BackColor = &H8000000F
Label12.BackColor = &H8000000F
Label13.BackColor = &H8000000F
Label14.BackColor = vbBlack
Label15.BackColor = &H8000000F
Label16.BackColor = &H8000000F
Label17.BackColor = &H8000000F
Label18.BackColor = vbBlack
Label19.BackColor = &H8000000F
Label20.BackColor = &H8000000F
Label21.BackColor = vbBlack
Label22.BackColor = vbBlack
Label23.BackColor = vbBlack
Label24.BackColor = vbBlack
Label25.BackColor = vbBlack
End If

If c = 3 Then
Label1.BackColor = &H8000000F
Label2.BackColor = vbBlack
Label3.BackColor = vbBlack
Label4.BackColor = vbBlack
Label5.BackColor = &H8000000F
Label6.BackColor = vbBlack
Label7.BackColor = &H8000000F
Label8.BackColor = &H8000000F
Label9.BackColor = &H8000000F
Label10.BackColor = vbBlack
Label11.BackColor = &H8000000F
Label12.BackColor = &H8000000F
Label13.BackColor = vbBlack
Label14.BackColor = vbBlack
Label15.BackColor = &H8000000F
Label16.BackColor = vbBlack
Label17.BackColor = &H8000000F
Label18.BackColor = &H8000000F
Label19.BackColor = &H8000000F
Label20.BackColor = vbBlack
Label21.BackColor = &H8000000F
Label22.BackColor = vbBlack
Label23.BackColor = vbBlack
Label24.BackColor = vbBlack
Label25.BackColor = &H8000000F
End If

If c = 4 Then
Label1.BackColor = &H8000000F
Label2.BackColor = &H8000000F
Label3.BackColor = vbBlack
Label4.BackColor = vbBlack
Label5.BackColor = &H8000000F
Label6.BackColor = &H8000000F
Label7.BackColor = vbBlack
Label8.BackColor = &H8000000F
Label9.BackColor = vbBlack
Label10.BackColor = &H8000000F
Label11.BackColor = vbBlack
Label12.BackColor = &H8000000F
Label13.BackColor = &H8000000F
Label14.BackColor = vbBlack
Label15.BackColor = &H8000000F
Label16.BackColor = vbBlack
Label17.BackColor = vbBlack
Label18.BackColor = vbBlack
Label19.BackColor = vbBlack
Label20.BackColor = vbBlack
Label21.BackColor = &H8000000F
Label22.BackColor = &H8000000F
Label23.BackColor = &H8000000F
Label24.BackColor = vbBlack
Label25.BackColor = &H8000000F
End If

If c = 5 Then
Label1.BackColor = vbBlack
Label2.BackColor = vbBlack
Label3.BackColor = vbBlack
Label4.BackColor = vbBlack
Label5.BackColor = vbBlack
Label6.BackColor = vbBlack
Label7.BackColor = &H8000000F
Label8.BackColor = &H8000000F
Label9.BackColor = &H8000000F
Label10.BackColor = &H8000000F
Label11.BackColor = vbBlack
Label12.BackColor = vbBlack
Label13.BackColor = vbBlack
Label14.BackColor = vbBlack
Label15.BackColor = &H8000000F
Label16.BackColor = &H8000000F
Label17.BackColor = &H8000000F
Label18.BackColor = &H8000000F
Label19.BackColor = &H8000000F
Label20.BackColor = vbBlack
Label21.BackColor = vbBlack
Label22.BackColor = vbBlack
Label23.BackColor = vbBlack
Label24.BackColor = vbBlack
Label25.BackColor = &H8000000F
End If

If c = 6 Then
Label1.BackColor = &H8000000F
Label2.BackColor = vbBlack
Label3.BackColor = vbBlack
Label4.BackColor = vbBlack
Label5.BackColor = &H8000000F
Label6.BackColor = vbBlack
Label7.BackColor = &H8000000F
Label8.BackColor = &H8000000F
Label9.BackColor = &H8000000F
Label10.BackColor = &H8000000F
Label11.BackColor = vbBlack
Label12.BackColor = vbBlack
Label13.BackColor = vbBlack
Label14.BackColor = vbBlack
Label15.BackColor = &H8000000F
Label16.BackColor = vbBlack
Label17.BackColor = &H8000000F
Label18.BackColor = &H8000000F
Label19.BackColor = &H8000000F
Label20.BackColor = vbBlack
Label21.BackColor = &H8000000F
Label22.BackColor = vbBlack
Label23.BackColor = vbBlack
Label24.BackColor = vbBlack
Label25.BackColor = &H8000000F
End If

If c = 7 Then
Label1.BackColor = vbBlack
Label2.BackColor = vbBlack
Label3.BackColor = vbBlack
Label4.BackColor = vbBlack
Label5.BackColor = vbBlack
Label6.BackColor = &H8000000F
Label7.BackColor = &H8000000F
Label8.BackColor = &H8000000F
Label9.BackColor = &H8000000F
Label10.BackColor = vbBlack
Label11.BackColor = &H8000000F
Label12.BackColor = &H8000000F
Label13.BackColor = &H8000000F
Label14.BackColor = vbBlack
Label15.BackColor = &H8000000F
Label16.BackColor = &H8000000F
Label17.BackColor = &H8000000F
Label18.BackColor = vbBlack
Label19.BackColor = &H8000000F
Label20.BackColor = &H8000000F
Label21.BackColor = &H8000000F
Label22.BackColor = vbBlack
Label23.BackColor = &H8000000F
Label24.BackColor = &H8000000F
Label25.BackColor = &H8000000F
End If

If c = 8 Then
Label1.BackColor = &H8000000F
Label2.BackColor = vbBlack
Label3.BackColor = vbBlack
Label4.BackColor = vbBlack
Label5.BackColor = &H8000000F
Label6.BackColor = vbBlack
Label7.BackColor = &H8000000F
Label8.BackColor = &H8000000F
Label9.BackColor = &H8000000F
Label10.BackColor = vbBlack
Label11.BackColor = &H8000000F
Label12.BackColor = vbBlack
Label13.BackColor = vbBlack
Label14.BackColor = vbBlack
Label15.BackColor = &H8000000F
Label16.BackColor = vbBlack
Label17.BackColor = &H8000000F
Label18.BackColor = &H8000000F
Label19.BackColor = &H8000000F
Label20.BackColor = vbBlack
Label21.BackColor = &H8000000F
Label22.BackColor = vbBlack
Label23.BackColor = vbBlack
Label24.BackColor = vbBlack
Label25.BackColor = &H8000000F
End If

If c = 9 Then
Label1.BackColor = &H8000000F
Label2.BackColor = vbBlack
Label3.BackColor = vbBlack
Label4.BackColor = vbBlack
Label5.BackColor = &H8000000F
Label6.BackColor = vbBlack
Label7.BackColor = &H8000000F
Label8.BackColor = &H8000000F
Label9.BackColor = &H8000000F
Label10.BackColor = vbBlack
Label11.BackColor = &H8000000F
Label12.BackColor = vbBlack
Label13.BackColor = vbBlack
Label14.BackColor = vbBlack
Label15.BackColor = vbBlack
Label16.BackColor = &H8000000F
Label17.BackColor = &H8000000F
Label18.BackColor = &H8000000F
Label19.BackColor = &H8000000F
Label20.BackColor = vbBlack
Label21.BackColor = vbBlack
Label22.BackColor = vbBlack
Label23.BackColor = vbBlack
Label24.BackColor = vbBlack
Label25.BackColor = &H8000000F
c = -1
End If
End Sub
