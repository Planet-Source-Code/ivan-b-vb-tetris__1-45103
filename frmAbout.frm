VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   2715
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1873.941
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1710
      Left            =   2940
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmAbout.frx":030A
      Top             =   101
      Width           =   2715
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2100
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "ivanbabic@vip.hr"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1050
      TabIndex        =   6
      Top             =   1575
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "e-mail:"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1050
      TabIndex        =   5
      Top             =   1365
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   165
      Picture         =   "frmAbout.frx":037D
      Top             =   300
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1304.511
      Y2              =   1304.511
   End
   Begin VB.Label lblDescription 
      Caption         =   "by Ivan Babic "
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1050
      TabIndex        =   1
      Top             =   1125
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   2
      Top             =   240
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1304.511
      Y2              =   1304.511
   End
   Begin VB.Label lblVersion 
      Height          =   225
      Left            =   1050
      TabIndex        =   3
      Top             =   780
      Width           =   1365
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

