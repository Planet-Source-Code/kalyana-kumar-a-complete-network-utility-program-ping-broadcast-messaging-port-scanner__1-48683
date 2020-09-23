VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPing 
   Caption         =   "Ping & Broadcast Messaging"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmPing.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox ProgressBG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7920
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   111
      Top             =   7800
      Visible         =   0   'False
      Width           =   3375
      Begin VB.PictureBox Progress 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   0
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   225
         TabIndex        =   113
         Top             =   0
         Width           =   3375
         Begin VB.Label lblProgress 
            Alignment       =   2  'Center
            BackColor       =   &H80000002&
            Caption         =   "Label3"
            ForeColor       =   &H80000009&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   114
            Top             =   0
            Width           =   3375
         End
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Label3"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   112
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ports"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9240
      TabIndex        =   18
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Map Setting"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9240
      TabIndex        =   17
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Non-Active Nodes"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9240
      TabIndex        =   16
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Active Nodes"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9240
      TabIndex        =   15
      Top             =   1440
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox IP3 
      ForeColor       =   &H80000006&
      Height          =   255
      Left            =   3600
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox IP4 
      ForeColor       =   &H80000006&
      Height          =   255
      Left            =   3960
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox IP2 
      ForeColor       =   &H80000006&
      Height          =   255
      Left            =   3240
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox IP1 
      ForeColor       =   &H80000006&
      Height          =   255
      Left            =   2880
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtFrom 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtMessage 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2880
      TabIndex        =   8
      Top             =   2520
      Width           =   3615
   End
   Begin VB.PictureBox pic16no 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   5880
      Picture         =   "frmPing.frx":000C
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   101
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic16yes 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   5880
      Picture         =   "frmPing.frx":00AA
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   100
      Top             =   6840
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox pic16 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5880
      Picture         =   "frmPing.frx":0147
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   99
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic19no 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   7680
      Picture         =   "frmPing.frx":02E7
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   97
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic19yes 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   7680
      Picture         =   "frmPing.frx":0385
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   96
      Top             =   6840
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox pic19 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7680
      Picture         =   "frmPing.frx":0422
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   95
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic17no 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   9480
      Picture         =   "frmPing.frx":05C2
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   93
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic17yes 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   9480
      Picture         =   "frmPing.frx":0660
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   92
      Top             =   6840
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox pic17 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   9480
      Picture         =   "frmPing.frx":06FD
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   91
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic18no 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   2160
      Picture         =   "frmPing.frx":089D
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   89
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic18yes 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2160
      Picture         =   "frmPing.frx":093B
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   88
      Top             =   6840
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox pic18 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2160
      Picture         =   "frmPing.frx":09D8
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   87
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic20no 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   4200
      Picture         =   "frmPing.frx":0B78
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   85
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic20yes 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4200
      Picture         =   "frmPing.frx":0C16
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   84
      Top             =   6840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pic20 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4200
      Picture         =   "frmPing.frx":0CB3
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   83
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic15yes 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2880
      Picture         =   "frmPing.frx":0E53
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   81
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic15no 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2880
      Picture         =   "frmPing.frx":0EF0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   80
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic15 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3120
      Picture         =   "frmPing.frx":0F8E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   79
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic14yes 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8640
      Picture         =   "frmPing.frx":112E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   77
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic14no 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8640
      Picture         =   "frmPing.frx":11CB
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   76
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic14 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8880
      Picture         =   "frmPing.frx":1269
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   75
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic13 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3600
      Picture         =   "frmPing.frx":1409
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   73
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic13yes 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3360
      Picture         =   "frmPing.frx":15A9
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   72
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic13no 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   3360
      Picture         =   "frmPing.frx":1646
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   71
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic12 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3960
      Picture         =   "frmPing.frx":16E4
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   69
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic12yes 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4200
      Picture         =   "frmPing.frx":1884
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   68
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic12no 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   4200
      Picture         =   "frmPing.frx":1921
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   67
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic11no 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   7680
      Picture         =   "frmPing.frx":19BF
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   65
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic11yes 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7680
      Picture         =   "frmPing.frx":1A5D
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   64
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic11 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7920
      Picture         =   "frmPing.frx":1AFA
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   63
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic10 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7080
      Picture         =   "frmPing.frx":1C9A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   61
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic10yes 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   6840
      Picture         =   "frmPing.frx":1E3A
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   60
      Top             =   5400
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox pic10no 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   6840
      Picture         =   "frmPing.frx":1ED7
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   59
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic9 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1560
      Picture         =   "frmPing.frx":1F75
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   57
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic9yes 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1320
      Picture         =   "frmPing.frx":2115
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   56
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic9no 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1320
      Picture         =   "frmPing.frx":21B2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   55
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic8no 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   8760
      Picture         =   "frmPing.frx":2250
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   54
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic8yes 
      BorderStyle     =   0  'None
      Height          =   200
      Left            =   8760
      Picture         =   "frmPing.frx":22EE
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   53
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic8 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   9000
      Picture         =   "frmPing.frx":238B
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   51
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic7yes 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2880
      Picture         =   "frmPing.frx":252B
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   49
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic7no 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2880
      Picture         =   "frmPing.frx":25C8
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   48
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic7 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2640
      Picture         =   "frmPing.frx":2666
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   47
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic6 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8400
      Picture         =   "frmPing.frx":2806
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   45
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic6no 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8160
      Picture         =   "frmPing.frx":29A6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   44
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic6yes 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8160
      Picture         =   "frmPing.frx":2A44
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   43
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtNoTimes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtInterval 
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.PictureBox pic5 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3240
      Picture         =   "frmPing.frx":2AE1
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   39
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic5no 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3480
      Picture         =   "frmPing.frx":2C81
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   38
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic5yes 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3480
      Picture         =   "frmPing.frx":2D1F
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   37
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic4yes 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8280
      Picture         =   "frmPing.frx":2DBC
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   35
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic4no 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8280
      Picture         =   "frmPing.frx":2E59
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   34
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic4 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8520
      Picture         =   "frmPing.frx":2EF7
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   33
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic3yes 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4680
      Picture         =   "frmPing.frx":3097
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   31
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic3no 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4680
      Picture         =   "frmPing.frx":3134
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   30
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic3 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4440
      Picture         =   "frmPing.frx":31D2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   29
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7320
      Picture         =   "frmPing.frx":3372
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   27
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic2no 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7080
      Picture         =   "frmPing.frx":3512
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   26
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic2yes 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7080
      Picture         =   "frmPing.frx":35B0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic1yes 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5880
      Picture         =   "frmPing.frx":364D
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   24
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic1no 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5880
      Picture         =   "frmPing.frx":36EA
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox MainNode 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5880
      Picture         =   "frmPing.frx":3788
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5880
      Picture         =   "frmPing.frx":3928
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdUnselect 
      Caption         =   "Unselect All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete Selected"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   1410
      Left            =   6840
      Style           =   1  'Checkbox
      TabIndex        =   19
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtFile 
      Height          =   195
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   150
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PING"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Insert a value and press ENTER"
      Height          =   255
      Left            =   4440
      TabIndex        =   115
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblInterval 
      Caption         =   "Interval"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   110
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5640
      TabIndex        =   109
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   108
      Top             =   7800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "Node"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6360
      TabIndex        =   107
      Top             =   7800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Scanning started ::: Scanning"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   106
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblhost 
      Caption         =   "HOST"
      Height          =   255
      Left            =   5760
      TabIndex        =   105
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   104
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   103
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lbl16 
      Height          =   255
      Left            =   5400
      TabIndex        =   102
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line38 
      Visible         =   0   'False
      X1              =   9600
      X2              =   9600
      Y1              =   6720
      Y2              =   6840
   End
   Begin VB.Label lbl19 
      Height          =   255
      Left            =   7200
      TabIndex        =   98
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line37 
      Visible         =   0   'False
      X1              =   7800
      X2              =   7800
      Y1              =   6720
      Y2              =   6840
   End
   Begin VB.Label lbl17 
      Height          =   255
      Left            =   9000
      TabIndex        =   94
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line36 
      Visible         =   0   'False
      X1              =   4320
      X2              =   4320
      Y1              =   6720
      Y2              =   6840
   End
   Begin VB.Label lbl18 
      Height          =   255
      Left            =   1560
      TabIndex        =   90
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line35 
      Visible         =   0   'False
      X1              =   2280
      X2              =   2280
      Y1              =   6720
      Y2              =   6840
   End
   Begin VB.Label lbl20 
      Height          =   255
      Left            =   3720
      TabIndex        =   86
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line34 
      Visible         =   0   'False
      X1              =   6000
      X2              =   6000
      Y1              =   5400
      Y2              =   6840
   End
   Begin VB.Line Line33 
      Visible         =   0   'False
      X1              =   6240
      X2              =   9600
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line32 
      Visible         =   0   'False
      X1              =   5760
      X2              =   2280
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line31 
      Visible         =   0   'False
      X1              =   5760
      X2              =   5760
      Y1              =   5400
      Y2              =   6720
   End
   Begin VB.Line Line30 
      Visible         =   0   'False
      X1              =   6240
      X2              =   6240
      Y1              =   5400
      Y2              =   6720
   End
   Begin VB.Label lbl15 
      Height          =   255
      Left            =   1560
      TabIndex        =   82
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line29 
      Visible         =   0   'False
      X1              =   3360
      X2              =   5520
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line28 
      Visible         =   0   'False
      X1              =   5520
      X2              =   5520
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Line27 
      Visible         =   0   'False
      X1              =   6480
      X2              =   6480
      Y1              =   5400
      Y2              =   6480
   End
   Begin VB.Line Line26 
      Visible         =   0   'False
      X1              =   6480
      X2              =   8640
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label lbl14 
      Height          =   255
      Left            =   9240
      TabIndex        =   78
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbl13 
      Height          =   255
      Left            =   3840
      TabIndex        =   74
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line25 
      Visible         =   0   'False
      X1              =   3480
      X2              =   3480
      Y1              =   5400
      Y2              =   5160
   End
   Begin VB.Label lbl12 
      Height          =   255
      Left            =   2520
      TabIndex        =   70
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbl11 
      Height          =   255
      Left            =   8280
      TabIndex        =   66
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line20 
      Visible         =   0   'False
      X1              =   7680
      X2              =   6600
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line23 
      Visible         =   0   'False
      X1              =   4440
      X2              =   5400
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line22 
      Visible         =   0   'False
      X1              =   5400
      X2              =   5400
      Y1              =   5400
      Y2              =   6120
   End
   Begin VB.Line Line19 
      Visible         =   0   'False
      X1              =   6600
      X2              =   6600
      Y1              =   5400
      Y2              =   6120
   End
   Begin VB.Label lbl10 
      Height          =   255
      Left            =   7320
      TabIndex        =   62
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbl9 
      Height          =   255
      Left            =   1800
      TabIndex        =   58
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbl8 
      Height          =   255
      Left            =   9240
      TabIndex        =   52
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line15 
      Visible         =   0   'False
      X1              =   8880
      X2              =   8880
      Y1              =   5160
      Y2              =   5400
   End
   Begin VB.Line Line18 
      Visible         =   0   'False
      X1              =   6960
      X2              =   6960
      Y1              =   5160
      Y2              =   5400
   End
   Begin VB.Line Line17 
      Visible         =   0   'False
      X1              =   1440
      X2              =   1440
      Y1              =   5160
      Y2              =   5400
   End
   Begin VB.Line Line14 
      Visible         =   0   'False
      X1              =   6600
      X2              =   8880
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line16 
      Visible         =   0   'False
      X1              =   5400
      X2              =   1440
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label lbl7 
      Height          =   255
      Left            =   1320
      TabIndex        =   50
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbl6 
      Height          =   255
      Left            =   8760
      TabIndex        =   46
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line13 
      Visible         =   0   'False
      X1              =   3120
      X2              =   5400
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line12 
      Visible         =   0   'False
      X1              =   6600
      X2              =   8160
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label4 
      Caption         =   "No of Times to Ping"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   42
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "[In Seconds]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   41
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lbl5 
      Height          =   255
      Left            =   1800
      TabIndex        =   40
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line10 
      Visible         =   0   'False
      X1              =   5520
      X2              =   5520
      Y1              =   4800
      Y2              =   4560
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00000000&
      Visible         =   0   'False
      X1              =   3720
      X2              =   5520
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line6 
      Visible         =   0   'False
      X1              =   6600
      X2              =   6960
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      Visible         =   0   'False
      X1              =   6960
      X2              =   8280
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line7 
      Visible         =   0   'False
      X1              =   6960
      X2              =   6960
      Y1              =   4680
      Y2              =   4440
   End
   Begin VB.Label lbl4 
      Height          =   255
      Left            =   8880
      TabIndex        =   36
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbl3 
      Height          =   255
      Left            =   3000
      TabIndex        =   32
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbl2 
      Height          =   255
      Left            =   7680
      TabIndex        =   28
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbl1 
      Height          =   255
      Left            =   6240
      TabIndex        =   23
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Visible         =   0   'False
      X1              =   6000
      X2              =   6000
      Y1              =   4680
      Y2              =   4080
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   6360
      X2              =   6360
      Y1              =   4680
      Y2              =   4080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      Visible         =   0   'False
      X1              =   6360
      X2              =   7080
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   5640
      X2              =   5640
      Y1              =   4680
      Y2              =   4080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      Visible         =   0   'False
      X1              =   4920
      X2              =   5640
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label1 
      Caption         =   "IP Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'* Developer: Kalyana Kumar aka Holymac                         *
'* Date Last Modified: Wednesday, September 17, 2000            *
'* Description: Ping program is a network utility to check the  *
'*              the internet/intranet connections of other      *
'*              computers in your office or home. This is       *
'*              for those people that don't like the            *
'*              command prompt and prefer the GUI.              *
'****************************************************************
Option Explicit

'**************** Broadcast Messaging's Syntax ********************************
'Standard declaration used in sending boradcast messaging
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_BAD_NETPATH As Long = 53
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const ERROR_NOT_SUPPORTED As Long = 50
Private Const ERROR_INVALID_NAME As Long = 123
Private Const NERR_BASE As Long = 2100
Private Const NERR_SUCCESS As Long = 0
Private Const NERR_NetworkError As Long = (NERR_BASE + 36)
Private Const NERR_NameNotFound As Long = (NERR_BASE + 173)
Private Const NERR_UseNotFound As Long = (NERR_BASE + 150)
'***********************
'method to get the current running version of Windows is to use the GetVersionEx API function
'Declaration for getting the Operating System's version
'All the declaration below are the declaration required to use this API
Private Declare Function GetVersionEx Lib "kernel32" _
Alias "GetVersionExA" _
(lpVersionInformation As OSVERSIONINFO) As Long

Private Const MAX_COMPUTERNAME As Long = 15
Private Const VER_PLATFORM_WIN32s As Long = 0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32_NT As Long = 2

Private Type OSVERSIONINFO
OSVSize         As Long
dwVerMajor      As Long
dwVerMinor      As Long
dwBuildNumber   As Long
PlatformID      As Long
szCSDVersion    As String * 128
End Type

'***********************

'User-defined type for passing the data to the Send function
Private Type NetMessageData
sServerName As String
sSendTo As String
sSendFrom As String
sMessage As String
End Type

'NetMessageBufferSend parameters:
'servername:  Unicode string specifying the name of the
'             remote server on which the function is to
'             execute. If this parameter is vbNullString,
'             the local computer is used.
'
'msgname:     Unicode string specifying the message alias to
'             which the message buffer should be sent.
'
'fromname:    Unicode string specifying who the message is from.
'             This parameter is required to send interrupting messages
'             from the computer name. If this parameter is NULL, the
'             message is sent from the logged-on user.
'
'msgbuf:      Unicode string containing the message to send.
'
'msgbuflen:   value that contains the length, in bytes, of
'             the message text pointed to by the msgbuf parameter.
Private Declare Function NetMessageBufferSend Lib "netapi32" _
(ByVal servername As String, _
ByVal msgname As String, _
ByVal fromname As String, _
ByVal msgbuf As String, _
ByRef msgbuflen As Long) As Long

'Declaration for getting the current computer name
Private Declare Function GetComputerName Lib "kernel32" _
Alias "GetComputerNameA" _
(ByVal lpBuffer As String, _
nSize As Long) As Long


'***************** Ping Syntax's ****************************************
'Declaration for using the delay function
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'General program's globl variable declaration
Dim reccount As Integer
Dim Status As String
Public Times As String
Public ActFlag As Integer
Public NotActFlag As Integer
Public Duplicate As Integer
Public SentMessage As Integer

'**************** Port Scanning Declaration *****************************
Public PortService As String
Public Explanation As String

'************ Declaration to enable to open a particular text file ******************************
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

'***********
Public Sub Pause(Duration As Single)
    Dim Current As Single
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
'This codes is used to select all the IP Address listed in the list box
Private Sub cmdSelect_Click()
    'General declaration
    Dim J As Integer
    'Forming the FOR loop
    'frmPing.List1.ListCount will give the total number of IP Address displayed in the listbox
    For J = 0 To frmPing.List1.ListCount - 1
    'J will identify which node in the list box. The first IP Address will have J=0.
    'In a listbox the list will start from 0
    'frmPing.List1.Selected(J)= FALSE means that particular node is not selected
    If frmPing.List1.Selected(J) = False Then
    'frmPing.List1.Selected(J)= TRUE means select that particular node in the list box
    frmPing.List1.Selected(J) = True
    End If
    Next J
End Sub
'This codes is used to un-select all the IP Address listed in the list box
Private Sub cmdUnselect_Click()
    'General declaration
    Dim J As Integer
    'Forming the FOR loop
    'frmPing.List1.ListCount will give the total number of IP Address displayed in the listbox
    For J = 0 To frmPing.List1.ListCount - 1
    'J will identify which node in the list box. The first IP Address will have J=0.
    'In a listbox the list will start from 0
    'frmPing.List1.Selected(J)= TRUE means select that particular node in the list box
    If frmPing.List1.Selected(J) = True Then
    'frmPing.List1.Selected(J)= FALSE means that particular node is not selected
    frmPing.List1.Selected(J) = False
    End If
    Next J
End Sub
'The sysntax executed when the PING command button is pressed
Private Sub Command1_Click()

'General variable declaration
ActFlag = 0
NotActFlag = 0
reccount = 1
Dim lkk As Integer
Dim strTemp As String
Dim Path As String

'Changing the mouse pointer
frmPing.MousePointer = 11

'Enabling the command button
Command1.Enabled = False
Command4.Enabled = False
cmdAdd.Enabled = False
cmdSelect.Enabled = False
cmdUnselect.Enabled = False
List1.Enabled = False

'Checking whethere the txtinterval is empty
If txtInterval.Text <> "" Then

'Checking whethere the txtnotimes is empty
If txtNoTimes.Text <> "" Then

    'Opening Active.txt and Not Active.txt
    Open App.Path & "\Active.txt" For Output As #10
    Close #10
    
    Open App.Path & "\Not Active.txt" For Output As #11
    Close #11
    'Opening Map Setting.txt for input
    Open App.Path & "\Map Setting.txt" For Input As #2
    'While the .txt is not end of file
    While Not EOF(2)
    'Read each line as strtemp
    Line Input #2, strTemp
    
    'Assign strtemp to Times
    Times = strTemp
    'The shell command used for ping the nodes
    'This command "Shell "command.com /c ping -n" is the core command
    'The Format: Shell Command & No of Times to ping & The Node[IP Address] > The file name to store the result
    Shell "command.com /c ping -n " & 2 & " " & Times & _
    " > c:\Ping.txt ", vbHide
    'Call function CheckConnection
    Call CheckConnection
    'Keep track of the reccount
    reccount = reccount + 1
              
    'Loop back
    Wend
    'Close map Setting.txt
    Close #2
    
    'Call function BroadcastMessaging
    Call BroadcastMessaging
Else
    'Display alert message
    MsgBox "Enter a value in the No of Times to Ping field!!"
    'Set the textboxes to empty
    txtNoTimes.Text = ""
    txtNoTimes.SetFocus
End If
Else
    'Display alert messages
    MsgBox "Enter a Interval!!"
    txtInterval.Text = ""
    txtInterval.SetFocus
End If
'Calling PortScanning function
Call PortScanning

'Reset certain fields or buttons
Command1.Enabled = True
'Setting the nouse pointer to their default state
frmPing.MousePointer = 0

Command2.Enabled = True
Command3.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command4.Enabled = True
cmdAdd.Enabled = True
cmdSelect.Enabled = True
cmdUnselect.Enabled = True
List1.Enabled = True
End Sub
'This is the function that checks what is the outcomeof the pinging process.
'When we ping a node, the outcome the ping is stored in Ping.txt
'ActFlag and NotActFlag is basically holds the value of 0 or 1
'ActFlag is used to identify active nodes
'NotActFlag is used to identify not active nodes
'For this 2 flags, when we verify the ping result we will set the values accordingly.Default is 0
Function CheckConnection()
'Local variable declaration
Dim Duration As Integer
Dim strLineData, strResult As String
Dim strResult2 As String
Dim strTemp As String
'Calculating the total interval duration
Duration = txtInterval.Text * 1000
'Delaying the program from continue executing for the set duration
Sleep Duration
txtFile.Text = ""

'Opening the log file for input as we are going to read content
Open "c:\Ping.txt" For Input As #1
'We will read line by line until it comes to end of file.
'And the resultis stored in a textbox
While Not EOF(1)
    'read  single line
    Line Input #1, strTemp
    'Concatenate the lines and assign the texts to txtfile textbox
    txtFile = txtFile & strTemp & vbCrLf
    'Loop again
    Wend
Close #1

'Check the texts in txtfile for a predetermined text using the INStr function
strResult = InStr(1, txtFile, "Reply from")
'If the result of Instr is not equal to 0 then the predetermined word does not exist.
If (strResult <> 0) Then
    'Act Flag is only a flag for us to use for identification
    If (ActFlag = 0) Then
        'Delete Active.txt
        Kill App.Path & "\Active.txt"
        'Open a new Active.txt for output
        'Opened as Output because fresh new entries are to entered
        'Existing Active.txt will be overwrittem
        'To check whether it is the first time entering data or more than ones
        'already, we use ActFlag as a counter
        Open App.Path & "\Active.txt" For Output As 4
        'Insert the IP Address into Active.txt
        Print #4, Times
        Close #4
        'Set status as active
        Status = "Active"
        'To check whether it is the first time entering data or more than ones
        'already, we use ActFlag as a counter
        'Set ActFlag = 1 because we already made the first entry.The next time
        'Active.txt is to be opened in Append mode
        ActFlag = 1
        'Call Logical Mapping funtion
        Call LogicalMap
    Else
        'Open a new Active.txt for output
        'Opened as Append so that we can keep on entering data without losing
        'existing data.
        Open App.Path & "\Active.txt" For Append As 4
        Print #4, Times
        Close #4
        Status = "Active"
        'To check whether it is the first time entering data or more than ones
        'already, we use ActFlag as a counter
        'Set ActFlag = 1 because we already made the first entry.The next time
        'Active.txt is to be opened in Append mode
        ActFlag = 1
        'Call Logical Map function
        Call LogicalMap
    End If
Else
    'If that particular nodes is not active
    If (NotActFlag = 0) Then
        'Why we need to kill and then open for output and the second time open for append.
        'The logic for Kill is to remove existing file and then create a new file. Why we need to that is because
        'existing file will hold the result of the previous ping so we need to clear that
        'Further more the first time we erite entry inside there,we need to open in in output mode
        'Output mode will remove all the contents of the file
        'The second time we write to the file in append mode
        'Append will write the details into the file without erasing the existing details and we can have
        'Continuity
        Kill App.Path & "\Not Active.txt"
        Open App.Path & "\Not Active.txt" For Output As 5
        Print #5, Times
        Close #5
        Status = "NotActive"
        NotActFlag = "1"
        Call LogicalMap
    Else
        Open App.Path & "\Not Active.txt" For Append As 5
        Print #5, Times
        Close #5
        Status = "NotActive"
        NotActFlag = 1
        Call LogicalMap
    End If
    
End If
End Function
'This module is used to scan all thenodes of the active nodes[1-1000]
'How it works: Open Active.txt and read the first node,scan the nodes until 1000 and the read the next node
'until finish.
'How it will check for the services:Only the ports from 1 to 1000 which are open will have a service running.
'If a particular node is closed then it will not have a service.
'When scanning,let say it detected port 80 as opened then straight away it will go to
'Private Sub Winsock1_Connect()- [winsock1 is the winsock components name,Connect is an event]
'So when a port is detected as open, the Connect event is executed
'Once it is executed, it will run the codes written inside and it will scan for the servics assigned to that
'port.
'If seen in _Connect module, all the services have been assigned because:
'The Well Known Ports are those from 0 through 1023 meaning each port have already been assigned a service from
'the beginning so we already know what are the services assigned for that particular port.
'So when a port is detected open then we go to Winsock1_Connect and then run through the IF stmts to check what
'services have been assigned to that port
Function PortScanning()
Dim strTemp As String
Dim I As Integer
'Opening Active.txt to get the active ports
Open App.Path & "\Active.txt" For Input As #39
    'Keep on looping
    While Not EOF(39)
    'Each line read is assigned to strtemp
    Line Input #39, strTemp
    
    'Opening Port.txt to enter some details
    Open App.Path & "\Port.txt" For Append As 35
    'Writing the required details in Port.txt
    Print #35, "IP Address :" & strTemp & "- Currently opened ports"
    Print #35, " "
    'Close the file
    Close #35
    
    Label2.Visible = True
    Label7.Visible = True
    Label8.Visible = True
    Label9.Visible = True
    'Progress Bar Initialization
    Me.ProgressBG.Visible = True
    
    'Text1.Text = strtemp
    Label7.Caption = strTemp
    'A for loop to mark the ports
    'Since we need to scan from 1-1000
    For I = 1 To 1000
    'Close any connection if it was already opened earlier
    Winsock1.Close
    'Assign the IP Address we need to connect which was read from Active.txt
    Winsock1.RemoteHost = strTemp
    'Assign which port of the IP Address to connect
    Winsock1.RemotePort = I
    'Label2.Caption = "Scanning Port ", it will show which port we are accessing
    Label9.Caption = I
    'Connect to the IP Address
    'Once this command is executed, it will try to connect to assigned port of that particular node
    'If Connect is successful then it will go to Winsock1_Connect to find out which service it is assigned to
    'If connect is not successful Winsock1_Connect will not be called but it will proceed to the next
    Winsock1.Connect
    'it must pause for about .006 of a second because it will be too much for the
    'computer to handle if you do any less
    Pause (0.006)
    'Progress Bar for Port scanning
    Me.Progress.Width = (Label9.Caption - 1) / (1000 - 1) * Me.ProgressBG.ScaleWidth
    Me.lblProgress(0).Caption = Int((Label9.Caption - 1) / (1000 - 1) * 100) & " %" '
    Me.lblProgress(1).Caption = Int((Label9.Caption - 1) / (1000 - 1) * 100) & " %"
    Next
    Wend
Close #39
   'Progress Bar Initialization
    Me.ProgressBG.Visible = False
End Function
'This module is used to add new IP Adddress to the Map Setting.txt
'It will check whether the entered IP Adddress is a already exist.If yes then it will alert the user
Private Sub cmdAdd_Click()
'General variable declaration
Dim IPAdd As String
Dim J As Integer

'Settinf default as 0
Duplicate = 0
cmdAdd.Enabled = False
'Joining all the numbers to form a valid IP Address
IPAdd = IP1.Text & "." & IP2.Text & "." & IP3.Text & "." & IP4.Text

    'Loop through the listbox to check whethere the newly to be entered IPAddress
    'already exist in the list
    'Basically to checl for duplication
    For J = 0 To frmPing.List1.ListCount - 1
        If frmPing.List1.List(J) = IPAdd Then
            'When looping through the list,it will chek whether the IP Address entered is the same as in the
            'listbox. If yes then it will set Duplicate = 1 meaning there us a same IP Address
            Duplicate = "1"
        Else
        End If
    Next J

'If Duplication exists then dispay alert messages
If Duplicate = "1" Then
    MsgBox "Duplication of IP Address are not allowed!"
    'Clear all the fields and ask user to enter new IP Address
    IP1.Text = ""
    IP2.Text = ""
    IP3.Text = ""
    IP4.Text = ""
    IP1.SetFocus
Else
    'If no duplication then..
    If (IP1.Text = "" Or IP2.Text = "" Or IP3.Text = "" Or IP4.Text = "") Then
        MsgBox "Please enter a IP Address [EG: 197.10.10.140]"
    Else
       'If duplication does not exist then add the IP Address to Map Setting.txt
        Open App.Path & "\Map Setting.txt" For Append As 3
        Print #3, IPAdd
        Close #3
        MsgBox "Node successfully added!"
    End If
    'Call LoadLIst function to reload all the IP Address so that the most updated
    'IP Address list can be displayed
    Call LoadList
    IP1.Text = ""
    IP2.Text = ""
    IP3.Text = ""
    IP4.Text = ""
    IP1.SetFocus
End If
End Sub
'This function will display the graphical output of the whole ping process
'All the nodes based on their ping result will be displayed in a graphical manner
'If a ping outcome for a particular node is active then i will be shown as ACTIVE.
'If a ping outcome for a particular node is notactive then i will be shown as NOTACTIVE.
Function LogicalMap()
    
'Reccount will keep track the number of the node whether it is the 1st or 5th node
'Status will be given when the ping outcome for each particular node is detemined
If (reccount = 1 And Status = "Active") Then
'Setting the line's borderstyle.Check the line's properties and you can know other available styles
Line1.BorderStyle = 1
'Set the line color
Line1.BorderColor = &HFF0000
'Assign the IP Address we are generating this diagram for
lbl1.Caption = Times
'Make the line and othert pics visible
lbl1.Visible = True
Line1.Visible = True
pic1.Visible = True
pic1yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 1 And Status = "NotActive") Then
Line1.BorderStyle = 5
Line1.BorderColor = &HFF&
lbl1.Caption = Times
lbl1.Visible = True
Line1.Visible = True
pic1.Visible = True
pic1no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'********
If (reccount = 2 And Status = "Active") Then
Line2.BorderStyle = 1
Line3.BorderStyle = 1
Line2.BorderColor = &HFF0000
Line3.BorderColor = &HFF0000
lbl2.Caption = Times
lbl2.Visible = True
Line2.Visible = True
Line3.Visible = True
pic2.Visible = True
pic2yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 2 And Status = "NotActive") Then
Line2.BorderStyle = 5
Line3.BorderStyle = 5
Line2.BorderColor = &HFF&
Line3.BorderColor = &HFF&
lbl2.Caption = Times
lbl2.Visible = True
Line2.Visible = True
Line3.Visible = True
pic2.Visible = True
pic2no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'********
If (reccount = 3 And Status = "Active") Then
Line4.BorderStyle = 1
Line5.BorderStyle = 1
Line4.BorderColor = &HFF0000
Line5.BorderColor = &HFF0000
lbl3.Caption = Times
lbl3.Visible = True
Line4.Visible = True
Line5.Visible = True
pic3.Visible = True
pic3yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 3 And Status = "NotActive") Then
Line4.BorderStyle = 5
Line5.BorderStyle = 5
Line4.BorderColor = &HFF&
Line5.BorderColor = &HFF&
lbl3.Caption = Times
lbl3.Visible = True
Line4.Visible = True
Line5.Visible = True
pic3.Visible = True
pic3no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'********
If (reccount = 4 And Status = "Active") Then
Line6.BorderStyle = 1
Line7.BorderStyle = 1
Line6.BorderStyle = 1
Line6.BorderColor = &HFF0000
Line7.BorderColor = &HFF0000
Line8.BorderColor = &HFF0000
lbl4.Caption = Times
lbl4.Visible = True
Line6.Visible = True
Line7.Visible = True
Line8.Visible = True
pic4.Visible = True
pic4yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 4 And Status = "NotActive") Then
Line6.BorderStyle = 5
Line7.BorderStyle = 5
Line8.BorderStyle = 5
Line6.BorderColor = &HFF&
Line7.BorderColor = &HFF&
Line8.BorderColor = &HFF&
lbl4.Caption = Times
lbl4.Visible = True
Line6.Visible = True
Line7.Visible = True
Line8.Visible = True
pic4.Visible = True
pic4no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'********
If (reccount = 5 And Status = "Active") Then
Line10.BorderStyle = 1
Line11.BorderStyle = 1
Line10.BorderColor = &HFF0000
Line11.BorderColor = &HFF0000
lbl5.Caption = Times
lbl5.Visible = True
Line10.Visible = True
Line11.Visible = True
pic5.Visible = True
pic5yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 5 And Status = "NotActive") Then
Line10.BorderStyle = 5
Line11.BorderStyle = 5
Line10.BorderColor = &HFF&
Line11.BorderColor = &HFF&
lbl5.Caption = Times
lbl5.Visible = True
Line10.Visible = True
Line11.Visible = True
pic5.Visible = True
pic5no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If
'********

If (reccount = 6 And Status = "Active") Then
Line12.BorderStyle = 1
Line12.BorderColor = &HFF0000
lbl6.Caption = Times
lbl6.Visible = True
Line12.Visible = True
pic6.Visible = True
pic6yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 6 And Status = "NotActive") Then
Line12.BorderStyle = 5
Line12.BorderColor = &HFF&
lbl6.Caption = Times
lbl6.Visible = True
Line12.Visible = True
pic6.Visible = True
pic6no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'*************
If (reccount = 7 And Status = "Active") Then
Line13.BorderStyle = 1
Line13.BorderColor = &HFF0000
lbl7.Caption = Times
lbl7.Visible = True
Line13.Visible = True
pic7.Visible = True
pic7yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 7 And Status = "NotActive") Then
Line13.BorderStyle = 5
Line13.BorderColor = &HFF&
lbl7.Caption = Times
lbl7.Visible = True
Line13.Visible = True
pic7.Visible = True
pic7no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'*************
If (reccount = 8 And Status = "Active") Then
Line14.BorderStyle = 1
Line15.BorderStyle = 1
Line14.BorderColor = &HFF0000
Line15.BorderColor = &HFF0000
lbl8.Caption = Times
lbl8.Visible = True
Line14.Visible = True
Line15.Visible = True
pic8.Visible = True
pic8yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 8 And Status = "NotActive") Then
Line14.BorderStyle = 5
Line15.BorderStyle = 5
Line14.BorderColor = &HFF&
Line15.BorderColor = &HFF&
lbl8.Caption = Times
lbl8.Visible = True
Line14.Visible = True
Line15.Visible = True
pic8.Visible = True
pic8no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'**************
If (reccount = 9 And Status = "Active") Then
Line16.BorderStyle = 1
Line17.BorderStyle = 1
Line16.BorderColor = &HFF0000
Line17.BorderColor = &HFF0000
lbl9.Caption = Times
lbl9.Visible = True
Line16.Visible = True
Line17.Visible = True
pic9.Visible = True
pic9yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 9 And Status = "NotActive") Then
Line16.BorderStyle = 5
Line17.BorderStyle = 5
Line16.BorderColor = &HFF&
Line17.BorderColor = &HFF&
lbl9.Caption = Times
lbl9.Visible = True
Line16.Visible = True
Line17.Visible = True
pic9.Visible = True
pic9no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'******************
If (reccount = 10 And Status = "Active") Then
Line18.BorderStyle = 1
Line18.BorderColor = &HFF0000
lbl10.Caption = Times
lbl10.Visible = True
Line18.Visible = True
pic10.Visible = True
pic10yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 10 And Status = "NotActive") Then
Line18.BorderStyle = 5
Line18.BorderColor = &HFF&
lbl10.Caption = Times
lbl10.Visible = True
Line18.Visible = True
pic10.Visible = True
pic10no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'*********

If (reccount = 11 And Status = "Active") Then
Line19.BorderStyle = 1
Line20.BorderStyle = 1
Line19.BorderColor = &HFF0000
Line20.BorderColor = &HFF0000
lbl11.Caption = Times
lbl11.Visible = True
Line19.Visible = True
Line20.Visible = True
pic11.Visible = True
pic11yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 11 And Status = "NotActive") Then
Line19.BorderStyle = 5
Line20.BorderStyle = 5
Line19.BorderColor = &HFF&
Line20.BorderColor = &HFF&
lbl11.Caption = Times
lbl11.Visible = True
Line19.Visible = True
Line20.Visible = True
pic11.Visible = True
pic11no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'**********
If (reccount = 12 And Status = "Active") Then
Line22.BorderStyle = 1
Line23.BorderStyle = 1
Line22.BorderColor = &HFF0000
Line23.BorderColor = &HFF0000
lbl12.Caption = Times
lbl12.Visible = True
Line22.Visible = True
Line23.Visible = True
pic12.Visible = True
pic12yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 12 And Status = "NotActive") Then
Line22.BorderStyle = 5
Line23.BorderStyle = 5
Line22.BorderColor = &HFF&
Line23.BorderColor = &HFF&
lbl12.Caption = Times
lbl12.Visible = True
Line22.Visible = True
Line23.Visible = True
pic12.Visible = True
pic12no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'****
If (reccount = 13 And Status = "Active") Then
Line25.BorderStyle = 1
Line25.BorderColor = &HFF0000
lbl13.Caption = Times
lbl13.Visible = True
Line25.Visible = True
pic13.Visible = True
pic13yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 13 And Status = "NotActive") Then
Line25.BorderStyle = 5
Line25.BorderColor = &HFF&
lbl13.Caption = Times
lbl13.Visible = True
Line25.Visible = True
pic13.Visible = True
pic13no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'********
If (reccount = 14 And Status = "Active") Then
Line26.BorderStyle = 1
Line27.BorderStyle = 1
Line26.BorderColor = &HFF0000
Line27.BorderColor = &HFF0000
lbl14.Caption = Times
lbl14.Visible = True
Line26.Visible = True
Line27.Visible = True
pic14.Visible = True
pic14yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 14 And Status = "NotActive") Then
Line26.BorderStyle = 5
Line27.BorderStyle = 5
Line26.BorderColor = &HFF&
Line27.BorderColor = &HFF&
lbl14.Caption = Times
lbl14.Visible = True
Line26.Visible = True
Line27.Visible = True
pic14.Visible = True
pic14no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'*******
If (reccount = 15 And Status = "Active") Then
Line28.BorderStyle = 1
Line29.BorderStyle = 1
Line28.BorderColor = &HFF0000
Line29.BorderColor = &HFF0000
lbl15.Caption = Times
lbl15.Visible = True
Line28.Visible = True
Line29.Visible = True
pic15.Visible = True
pic15yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 15 And Status = "NotActive") Then
Line28.BorderStyle = 5
Line29.BorderStyle = 5
Line28.BorderColor = &HFF&
Line29.BorderColor = &HFF&
lbl15.Caption = Times
lbl15.Visible = True
Line28.Visible = True
Line29.Visible = True
pic15.Visible = True
pic15no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'*********
If (reccount = 16 And Status = "Active") Then
Line34.BorderStyle = 1
Line34.BorderColor = &HFF0000
lbl16.Caption = Times
lbl16.Visible = True
Line34.Visible = True
pic16.Visible = True
pic16yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 16 And Status = "NotActive") Then
Line34.BorderStyle = 5
Line34.BorderColor = &HFF&
lbl16.Caption = Times
lbl16.Visible = True
Line34.Visible = True
pic16.Visible = True
pic16no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'*****
If (reccount = 17 And Status = "Active") Then
Line30.BorderStyle = 1
Line33.BorderStyle = 1
Line38.BorderStyle = 1
Line30.BorderColor = &HFF0000
Line33.BorderColor = &HFF0000
Line38.BorderColor = &HFF0000
lbl17.Caption = Times
lbl17.Visible = True
Line30.Visible = True
Line33.Visible = True
Line38.Visible = True
pic17.Visible = True
pic17yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 17 And Status = "NotActive") Then
Line30.BorderStyle = 5
Line33.BorderStyle = 5
Line38.BorderStyle = 5
Line30.BorderColor = &HFF&
Line33.BorderColor = &HFF&
Line38.BorderColor = &HFF&
lbl17.Caption = Times
lbl17.Visible = True
Line30.Visible = True
Line33.Visible = True
Line38.Visible = True
pic17.Visible = True
pic17no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'**********

If (reccount = 18 And Status = "Active") Then
Line31.BorderStyle = 1
Line32.BorderStyle = 1
Line35.BorderStyle = 1
Line31.BorderColor = &HFF0000
Line32.BorderColor = &HFF0000
Line35.BorderColor = &HFF0000
lbl18.Caption = Times
lbl18.Visible = True
Line31.Visible = True
Line32.Visible = True
Line35.Visible = True
pic18.Visible = True
pic18yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 18 And Status = "NotActive") Then
Line31.BorderStyle = 5
Line32.BorderStyle = 5
Line35.BorderStyle = 5
Line31.BorderColor = &HFF&
Line32.BorderColor = &HFF&
Line35.BorderColor = &HFF&
lbl18.Caption = Times
lbl18.Visible = True
Line31.Visible = True
Line32.Visible = True
Line35.Visible = True
pic18.Visible = True
pic18no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If

'*********
If (reccount = 19 And Status = "Active") Then
Line37.BorderStyle = 1
Line37.BorderColor = &HFF0000
lbl19.Caption = Times
lbl19.Visible = True
Line37.Visible = True
pic19.Visible = True
pic19yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 19 And Status = "NotActive") Then
Line37.BorderStyle = 5
Line37.BorderColor = &HFF&
lbl19.Caption = Times
lbl19.Visible = True
Line37.Visible = True
pic19.Visible = True
pic19no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If
'*************
If (reccount = 20 And Status = "Active") Then
Line36.BorderStyle = 1
Line36.BorderColor = &HFF0000
lbl20.Caption = Times
lbl20.Visible = True
Line36.Visible = True
pic20.Visible = True
pic20yes.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
If (reccount = 20 And Status = "NotActive") Then
Line36.BorderStyle = 5
Line36.BorderColor = &HFF&
lbl20.Caption = Times
lbl20.Visible = True
Line36.Visible = True
pic20.Visible = True
pic20no.Visible = True
MainNode.Visible = True
lblhost.Visible = True
Else
End If
End If


End Function
'This function will be used to structure the message sending process and then send
'them to the message sending function
Function BroadcastMessaging()
'Variable declaration
Dim msgData As NetMessageData
Dim sSuccess As String
Dim strTemp As String

'OPen Active.txt for input
'This file will contain all the nodes which were verified to be active
Open App.Path & "\Active.txt" For Input As #22
While Not EOF(22)
Line Input #22, strTemp

'A message sending structure for a particular node
With msgData
'assign the Node the message to be sent[recipient]
.sSendTo = strTemp
'assign from whom the message is from
.sSendFrom = txtFrom.Text
'the message to be sent
.sMessage = txtMessage.Text
End With

'Once the structure have been updated, send the whole structure togetger with the assigned details like
'Message,From etc to NetSendMessage function - passing the msgData to the funtion
sSuccess = NetSendMessage(msgData)
'Perform the process as mentioned above for every single node deemed active
Wend
Close #22

Label1.Caption = sSuccess
End Function

Private Sub Command2_Click()
ShellExecute Me.hwnd, "open", App.Path & "\Active.txt", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub Command3_Click()
ShellExecute Me.hwnd, "open", App.Path & "\Not Active.txt", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

'This sub is executed when all the selected node need to be deleted
Private Sub Command4_Click()

Dim J, M As Integer
Dim TotalDel As Integer
Dim TotalToDel As Integer
TotalDel = 0
For M = 0 To frmPing.List1.ListCount - 1
If frmPing.List1.Selected(M) = False Then
    TotalDel = TotalDel + 1
    Close #3
End If
Next M

For M = 0 To frmPing.List1.ListCount - 1
If frmPing.List1.Selected(M) = True Then
    TotalToDel = TotalToDel + 1
    Close #3
End If
Next M

If (TotalDel = frmPing.List1.ListCount) Then
MsgBox "Please select the nodes to be deleted before pressing DELETE"
Else
Kill App.Path & "\Map Setting.txt"
'Go through the list to confirm which node are to be deleted and which are not.
For J = 0 To frmPing.List1.ListCount - 1
If frmPing.List1.Selected(J) = False Then
Open App.Path & "\Map Setting.txt" For Append As 3
  Print #3, List1.List(J)
    Close #3
End If
Next J

If (TotalToDel = frmPing.List1.ListCount) Then
Open App.Path & "\Map Setting.txt" For Output As #50
Close #50
Call LoadList
Else
Call LoadList
End If
MsgBox "Node successfully deleted!"
'Unload Me
'frmPing.Show
End If
End Sub

Private Sub Command5_Click()
ShellExecute Me.hwnd, "open", App.Path & "\Map Setting.txt", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub Command6_Click()
ShellExecute Me.hwnd, "open", App.Path & "\Port.txt", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub Form_Activate()
Dim retval As String

'Check the said file is exist at the said path

retval = Dir$("c:\Ping.txt")
'If the file exist then the file name will be passed back to retval
If retval = "Ping.txt" Then
Else
'if retval ="" which means the file does not exist than create a instance of that file
Open "c:\Ping.txt" For Output As #25
Close #25
End If

retval = Dir$(App.Path & "\Active.txt")

If retval = "Active.txt" Then
Else
Open App.Path & "\Active.txt" For Output As #26
Close #26
End If

retval = Dir$(App.Path & "\Not Active.txt")

If retval = "Not Active.txt" Then
Else
Open App.Path & "\Not Active.txt" For Output As #26
Close #26
End If

retval = Dir$(App.Path & "\Map Setting.txt")

If retval = "Map Setting.txt" Then
Else
Open App.Path & "\Map Setting.txt" For Output As #26
Close #26
End If

Open App.Path & "\Port.txt" For Output As #40
Close #40

Call LoadList
txtInterval.SetFocus
End Sub
'To list all the nodes in the list box
Function LoadList()
Dim count As Integer
Dim strTemp As String
List1.Clear
ActFlag = 0
NotActFlag = 0
'Open the file for input
Open App.Path & "\Map Setting.txt" For Input As #2
While Not EOF(2)
  'Read the first line in the txt file and then assign it to the list box
  Line Input #2, strTemp
   List1.AddItem strTemp
'Loop back.This will be looped until the end of the txt file
Wend
Close #2
End Function

Private Sub IP1_Change()
'Validation: Whenever something is typed in IP1 then it wil. call this module
'Call NumFilter(IP1, Val(IP1.Text))
If (Len(IP1) = 3) Then
IP2.SetFocus
Else
End If
End Sub

Private Sub IP1_KeyPress(KeyAscii As Integer)
'If ENTER is pressed
If (KeyAscii = "13") Then
If (IP1.Text <> "") Then
IP2.SetFocus
Else
MsgBox "Enter a valid value!!"
IP1.SetFocus
End If
Else
End If
End Sub

Private Sub IP2_Change()
'Validation
'Call NumFilter(IP2, Val(IP2.Text))
If (Len(IP2) = 3) Then
IP3.SetFocus
Else
End If
End Sub

Private Sub IP2_KeyPress(KeyAscii As Integer)
If (KeyAscii = "13") Then
If (IP2.Text <> "") Then
IP3.SetFocus
Else
MsgBox "Enter a valid value!!"
IP2.SetFocus
End If
Else
End If
End Sub

Private Sub IP3_Change()
'Validation
'Call NumFilter(IP3, Val(IP3.Text))
If (Len(IP3) = 3) Then
IP4.SetFocus
Else
End If
End Sub

Private Sub IP3_KeyPress(KeyAscii As Integer)
If (KeyAscii = "13") Then
If (IP3.Text <> "") Then
IP4.SetFocus
Else
MsgBox "Enter a valid value!!"
IP4.SetFocus
End If
Else
End If
End Sub

Private Sub IP4_Change()
'Validation
'Call NumFilter(IP4, Val(IP4.Text))
cmdAdd.Enabled = True
End Sub

Private Sub IP4_KeyPress(KeyAscii As Integer)
If (KeyAscii = "13") Then
If (IP1.Text <> "" Or IP2.Text <> "" Or IP3.Text <> "" Or IP4.Text <> "") Then
cmdAdd_Click
Else
MsgBox "Enter a valid IP Address!!"
End If
Else
End If
End Sub

Private Sub List1_Click()
Dim Top As Integer
Top = List1.ListCount
End Sub

Private Sub txt1_Change()
Call NumFilter(txt1, Val(txt1.Text))
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
If (KeyAscii = "13") Then
If (txt1.Text <> "") Then
cmdAdd.SetFocus
Else
MsgBox "Enter a valid IP Address!!"
End If
Else
End If
End Sub
Private Sub txtFrom_KeyPress(KeyAscii As Integer)
If (KeyAscii = "13") Then
If (txtFrom.Text <> "") Then
Command1_Click
Else
End If
Else
End If
Command1.Enabled = True
End Sub

Private Sub txtInterval_Change()
Call NumFilter(txtInterval, Val(txtInterval.Text))
End Sub

Private Sub txtInterval_KeyPress(KeyAscii As Integer)
If (KeyAscii = "13") Then
If (txtInterval <> "") Then
IP1.Text = ""
IP2.Text = ""
IP3.Text = ""
IP4.Text = ""
txtNoTimes.Enabled = True
txtNoTimes.SetFocus
Else
MsgBox "Enter a Interval!!"
End If
Else
End If
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
If (KeyAscii = "13") Then
If (txtMessage.Text <> "") Then
txtFrom.Enabled = True
txtFrom.SetFocus
Else
MsgBox "Enter the Broadcast Message!!"
End If
Else
End If
End Sub

Private Sub txtNoTimes_Change()
Call NumFilter(txtNoTimes, Val(txtNoTimes.Text))
End Sub

Private Sub txtNoTimes_KeyPress(KeyAscii As Integer)
If (KeyAscii = "13") Then
If (txtNoTimes.Text <> "") Then
txtMessage.Enabled = True
txtMessage.SetFocus
Else
MsgBox "Enter a value in the No of Times to Ping field!!"
End If
Else
End If
End Sub
'Checks the Operatin Systems version
'This module uses API library and the declaration is shown above
Private Function IsWinNT() As Boolean

'returns True if running WinNT/Win2000/WinXP. This broadcast messaging will only work for NT/2000/XP or above
'If the windows are WIN32 versions then
#If Win32 Then

'General variable declaration
'OSVERSIONINFO will hold the OS's version details
'And here we are assigning the os version details to OSV
Dim OSV As OSVERSIONINFO

'OSV will hold the result
OSV.OSVSize = Len(OSV)

'If there exist a version for an operating system in WIN32 then we assume it is either WinNT/2000/WinXP
If GetVersionEx(OSV) = 1 Then

  'PlatformId contains a value representing the OS.
   IsWinNT = (OSV.PlatformID = VER_PLATFORM_WIN32_NT)
   
End If

#End If

End Function
Private Function NetSendMessage(msgData As NetMessageData) As String

Dim success As Long

'assure that the OS is NT ..
'NetMessageBufferSend  can not
'be called on Win9x
If IsWinNT() Then

With msgData

  'if To name omitted return error and exit
   If .sSendTo = "" Then
      'If sSendTo is empty then assign NetSendMessage as the equivalent of
      'GetNetSendMessageStatus(ERROR_INVALID_PARAMETER)
      'GetNetSendMessageStatus is a function
      'ERROR_INVALID_PARAMETER is a type of error
      'In GetNetSendMessageStatus we will check what type of error and based on that error type we will assign
      'a error message. To do that we need to call the function and send the error type
      NetSendMessage = GetNetSendMessageStatus(ERROR_INVALID_PARAMETER)
      Exit Function
      
   Else
 
     'if there is a message
      If Len(.sMessage) Then

        'convert the strings to unicode
         .sSendTo = StrConv(.sSendTo, vbUnicode)
         .sMessage = StrConv(.sMessage, vbUnicode)
      
        'Note that the API could be called passing
        'vbNullString as the SendFrom and sServerName
        'strings. This would generate the message on
        'the sending machine.
         If Len(.sServerName) > 0 Then
                'Setting .sServerName to the .sServerName converted into Unicode
                'StrConv is the function used to convert from 1 datatype to another
               .sServerName = StrConv(.sServerName, vbUnicode)
               'If the length is not > 0 then set it to empty
         Else: .sServerName = vbNullString
         End If
                  
         If Len(.sSendFrom) > 0 Then
                'Setting .sSendFrom to the .sSendFrom converted into Unicode
               .sSendFrom = StrConv(.sSendFrom, vbUnicode)
               'If the length is not > 0 then set it to empty
         Else: .sSendFrom = vbNullString
         End If
      
        'change the cursor and show. Control won't return
        'until the call has completed.
         Screen.MousePointer = vbHourglass
     
        'Call function NetMessageBufferSend
         success = NetMessageBufferSend(.sServerName, _
                                        .sSendTo, _
                                        .sSendFrom, _
                                        .sMessage, _
                                        ByVal Len(.sMessage))
     
         Screen.MousePointer = vbNormal
     
         NetSendMessage = GetNetSendMessageStatus(success)

      End If 'If Len(.sMessage)
   End If  'If .sSendTo
End With  'With msgData
End If  'If IsWinNT

End Function
'This are all the possible error that could happen when sending the messages
'In the error trapping, if a error occurs, those errors will be trapped and
'will be identifed and a understandable message will be shown
'All the Case error names are real error names generated by the system.
'We eill filter that names and will give back user a more meaningful and understandable message
Private Function GetNetSendMessageStatus(nError As Long) As String

Dim msg As String

Select Case nError
'Error Names                  'Our own intepreted error message based on the system generated errors
Case NERR_SUCCESS:            msg = "The message was successfully sent"
Case NERR_NameNotFound:       msg = "Send To not found"
Case NERR_NetworkError:       msg = "General network error occurred"
Case NERR_UseNotFound:        msg = "Network connection not found"
Case ERROR_ACCESS_DENIED:     msg = "Access to computer denied"
Case ERROR_BAD_NETPATH:       msg = "Sent From server name not found."
Case ERROR_INVALID_PARAMETER: msg = "Invalid parameter(s) specified."
Case ERROR_NOT_SUPPORTED:     msg = "Network request not supported."
Case ERROR_INVALID_NAME:      msg = "Illegal character or malformed name."
Case Else:                    msg = "Unknown error executing command."

End Select

GetNetSendMessageStatus = msg

End Function
Private Function TrimNull(item As String)

'return string before the terminating null
Dim pos As Integer

pos = InStr(item, Chr$(0))

If pos Then
   TrimNull = Left$(item, pos - 1)
Else: TrimNull = item
End If

End Function
'This function will take the values passed to it and check wheteher they are numbers
'or not.If they are not number  then they will be deleted or not displayed
Private Sub NumFilter(T As TextBox, ByVal I As Double)
On Error Resume Next

If I <> 0 Then
T.BackColor = &H80000005
T.Text = I
T.SelStart = Len(T)
Else
'Used to delete the  character than were entered if they are not integers
If Len(T) > 0 Then _
T.Text = Right(T.Text, Len(T) - 1)
T.SelStart = Len(T)
T.BackColor = &HFFF&
End If

End Sub
'This is the module which will be called in between when PortScanning module is been executed
'When ever in PortScanning, it detected a open port,this module will then be executed before going to the next
'port
Private Sub Winsock1_Connect()
'telling it to print what port if it is connected to if the connection is succesfull
Dim PortNo As Integer
PortNo = Label9.Caption

'If you see, we have assigned the services hardcoded.The reason being is because Port no is divided into 3
'categories
'The Well Known Ports are those from 0 through 1023.
'The Registered Ports are those from 1024 through 49151
'The Dynamic and/or Private Ports are those from 49152 through 65535
'Well known ports are the one they have assigned a service indefinitely.Meaning it will never change.
'EG If the port is 80 then definitely the service running in port 80 is HTTP
'Since we already know what are the nodes, it is easier for us to assign like this
If (PortNo = 0) Then
PortService = "Reserved"
Explanation = "This port is reserved"
End If

If (PortNo = 1) Then
PortService = "TCPMUX"
Explanation = "TCP Port Service Multiplexer"
End If

If (PortNo = 2) Then
PortService = "COMPRESSNET"
Explanation = "Management Utility"
End If

If (PortNo = 3) Then
PortService = "COMPRESSNET"
Explanation = "Compression Process"
End If
'********* UNASSIGNED PORTS ***************************
'Some of the ports from 1 - 1023 are still unassigned
If (PortNo = 4 Or PortNo = 8 Or PortNo = 10 Or PortNo = 12 Or PortNo = 14 Or PortNo = 15 Or PortNo = 16 Or PortNo = 26 Or PortNo = 28 Or PortNo = 30 _
    Or PortNo = 32 Or PortNo = 34 Or PortNo = 36 Or PortNo = 40 Or PortNo = 60 Or PortNo = 285 Or PortNo = 708 Or PortNo = 743 Or PortNo = 745 _
    Or PortNo = 746 Or PortNo = 755 Or PortNo = 756 Or PortNo = 766 Or PortNo = 768 Or PortNo = 778 Or PortNo = 779) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 268 And PortNo < 280) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 287 And PortNo < 308) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 322 And PortNo < 333) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 333 And PortNo < 344) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 699 And PortNo < 704) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 711 And PortNo < 729) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 731 And PortNo < 741) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 780 And PortNo < 786) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 785 And PortNo < 800) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 801 And PortNo < 810) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 810 And PortNo < 828) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 829 And PortNo < 847) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 848 And PortNo < 860) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 860 And PortNo < 872) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If


If (PortNo > 873 And PortNo < 886) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 888 And PortNo < 900) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 903 And PortNo < 911) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If

If (PortNo > 913 And PortNo < 989) Then
PortService = "UNASSIGNED"
Explanation = "Unassigned Port"
End If
'********* END OF UNASSIGNED PORTS ***************************
If (PortNo = 7) Then
PortService = "ECHO"
Explanation = "ECHO"
End If

If (PortNo = 9) Then
PortService = "DISCARD"
Explanation = "Discarded Port"
End If

If (PortNo = 11) Then
PortService = "SYSTAT"
Explanation = "Active Users"
End If

If (PortNo = 13) Then
PortService = "DAYTIME"
Explanation = "Daytime - RFC 867"
End If

If (PortNo = 17) Then
PortService = "QOTD"
Explanation = "Quote of the Day"
End If

If (PortNo = 18) Then
PortService = "MSP"
Explanation = "Message send Protocol"
End If

If (PortNo = 19) Then
PortService = "CHARGEN"
Explanation = "Character Generator"
End If

If (PortNo = 20 Or PortNo = 21) Then
PortService = "FTP/FTP-DATA"
Explanation = "File Transfer"
End If

If (PortNo = 22) Then
PortService = "SSH"
Explanation = "SSH Remote Login Protocol"
End If

If (PortNo = 23) Then
PortService = "TELNET"
Explanation = "Telnet"
End If

If (PortNo = 24) Then
PortService = "MAIL"
Explanation = "Any Private Mail System"
End If

If (PortNo = 25) Then
PortService = "SMTP"
Explanation = "Simple Mail Transfer"
End If

If (PortNo = 27) Then
PortService = "NSW-FE"
Explanation = "NSW User System FE"
End If

If (PortNo = 29) Then
PortService = "MSG-ICP"
Explanation = "MSG-ICP"
End If

If (PortNo = 31) Then
PortService = "MSG-AUTH"
Explanation = "MSG-Authentication"
End If

If (PortNo = 33) Then
PortService = "DSP"
Explanation = "Display Support Protocol"
End If

If (PortNo = 38) Then
PortService = "RAP"
Explanation = "Route Access Protocol"
End If

If (PortNo = 39) Then
PortService = "RLP"
Explanation = "Resource Location Protocol"
End If

If (PortNo = 41) Then
PortService = "Graphics"
Explanation = "Graphics"
End If

If (PortNo = 42) Then
PortService = "NameServer"
Explanation = "Host Name Server"
End If

If (PortNo = 43) Then
PortService = "NICKNAME"
Explanation = "Who Is"
End If

If (PortNo = 44) Then
PortService = "MPM-FLAGS"
Explanation = "MPM FLAGS Protocol"
End If

If (PortNo = 47) Then
PortService = "NI-FTP"
Explanation = "NI FTP"
End If

If (PortNo = 48) Then
PortService = "AUDITD"
Explanation = "Digital Audit Daemon"
End If

If (PortNo = 49) Then
PortService = "TACACS"
Explanation = "Login Host Protocol"
End If

If (PortNo = 50) Then
PortService = "RE-MAIL-CK"
Explanation = "Remote Mail Checking Protocol"
End If

If (PortNo = 51) Then
PortService = "LA-MAINT"
Explanation = "IMP Logical Address Maintenance"
End If

If (PortNo = 53) Then
PortService = "DOMAIN"
Explanation = "Domain Name Server"
End If

If (PortNo = 63) Then
PortService = "WHOIS++"
Explanation = "WHOIS++"
End If

If (PortNo = 64) Then
PortService = "COVIA"
Explanation = "Communications Integrator-CI"
End If

If (PortNo = 66) Then
PortService = "SQL*NET"
Explanation = "Oracle SQL*NET"
End If

If (PortNo = 67 Or PortNo = 68) Then
PortService = "BOOTPS"
Explanation = "Bootstrap Protocol Server/Client"
End If

If (PortNo = 70) Then
PortService = "GOPHER"
Explanation = "Gopher"
End If

If (PortNo = 80) Then
PortService = "HTTP/WWW/WWW-HTTP"
Explanation = "World Wide Web HTTP"
End If

If (PortNo = 92) Then
PortService = "NPP"
Explanation = "Network Printing Protocol"
End If

If (PortNo = 93) Then
PortService = "DCP"
Explanation = "Device Control Protocol"
End If

If (PortNo = 101) Then
PortService = "HOSTNAME"
Explanation = "NIC Host Name Server"
End If

If (PortNo = 105) Then
PortService = "CSNET-NS"
Explanation = "MailBox Name Server"
End If

If (PortNo = 107) Then
PortService = "RTELNET"
Explanation = "Remote Telnet Service"
End If

If (PortNo = 110) Then
PortService = "POP3"
Explanation = "Post Office Protocol - Version 3"
End If

If (PortNo = 111) Then
PortService = "SUNRPC"
Explanation = "SUN Remote Procedure Call"
End If

If (PortNo = 113) Then
PortService = "AUTH"
Explanation = "Authentication Service"
End If

If (PortNo = 115) Then
PortService = "SFTP"
Explanation = "Simple File Transfer Protocol"
End If

If (PortNo = 118) Then
PortService = "SQLSERV"
Explanation = "SQL Services"
End If

If (PortNo = 119) Then
PortService = "NNTP"
Explanation = "Network News Transfer Protocol"
End If

If (PortNo = 130 Or PortNo = 131 Or PortNo = 132) Then
PortService = "CISCO-FNA/TNA/SYS"
Explanation = "CISCO -FNative/TNative/SYSMAINT"
End If

If (PortNo = 135) Then
PortService = "EPMAP"
Explanation = "DCE EndPoint Resolution"
End If

If (PortNo >= 137 And PortNo <= 139) Then
PortService = "NETBIOS-SSN"
Explanation = "Netbios Session Service"
End If

If (PortNo = 143) Then
PortService = "IMAP"
Explanation = "Internet Messaging Access Protocol"
End If

If (PortNo = 158) Then
PortService = "PCMAIL-SRV"
Explanation = "PCMail Server"
End If

If (PortNo = 170) Then
PortService = "PRINT-SRV"
Explanation = "Network PostScript"
End If

If (PortNo = 171) Then
PortService = "MULTIPLEX"
Explanation = "Network Innovations Multiplex"
End If

If (PortNo = 179) Then
PortService = "BGP"
Explanation = "Border Gateway Protocol"
End If

If (PortNo = 189) Then
PortService = "QFT"
Explanation = "Queued File Transport"
End If

If (PortNo = 194) Then
PortService = "IRC"
Explanation = "Internet Relay Chat Protocol"
End If

If (PortNo = 200) Then
PortService = "SRC"
Explanation = "IBM System Resource Controller"
End If

If (PortNo = 217) Then
PortService = "DBASE"
Explanation = "dBASE Unix"
End If

If (PortNo = 220) Then
PortService = "IMPA3"
Explanation = "Interactive Mail Access Protocol v3"
End If

If (PortNo = 280) Then
PortService = "HTTP-MGMT"
Explanation = "HTTP-MGMT"
End If

If (PortNo = 359) Then
PortService = "NSRMP"
Explanation = "Network Security Risk Management Protocol"
End If

If (PortNo = 372) Then
PortService = "ULISTPROC"
Explanation = "ListProcessor"
End If

If (PortNo = 384) Then
PortService = "ARNS"
Explanation = "A Remote Network Server System"
End If

If (PortNo = 385) Then
PortService = "IBM-APP"
Explanation = "IBM Application"
End If

If (PortNo = 396) Then
PortService = "NETWARE-IP"
Explanation = "Novell Netware over IP"
End If

If (PortNo = 397) Then
PortService = "MPTN"
Explanation = "Multi Protocol Trans. NET"
End If

If (PortNo = 401) Then
PortService = "UPS"
Explanation = "Uninterruptable Power Supply"
End If

If (PortNo = 427) Then
PortService = "SVRLOC"
Explanation = "Server Location"
End If


If (PortNo = 443) Then
PortService = "HTTPS"
Explanation = "HTTP protocol over TLS/SSL"
End If

If (PortNo = 445) Then
PortService = "MICROSOFT-DS"
Explanation = "Microsoft-DS"
End If

If (PortNo = 449) Then
PortService = "AS-SERVERMAP"
Explanation = "AS Server Mapper"
End If

If (PortNo = 514) Then
PortService = "SHELL/SYSLOG"
Explanation = "CMD like EXEC"
End If

If (PortNo = 515) Then
PortService = "PRINTER"
Explanation = "Spooler"
End If

'Open Port.txt for append
'Insert all the details in the file
Open App.Path & "\Port.txt" For Append As 36
    Print #36, "Port Number     :" & " " & PortNo
    Print #36, "Keyword         :" & " " & PortService
    Print #36, "Description     :" & " " & Explanation
    Print #36, "  "
Close #36

End Sub

