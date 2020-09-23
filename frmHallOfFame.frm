VERSION 5.00
Begin VB.Form frmHallOfFame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Power Tetris - Hall Of Fame"
   ClientHeight    =   1635
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3105
   Icon            =   "frmHallOfFame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "II"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   300
      TabIndex        =   8
      Top             =   600
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "III"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   315
      Left            =   300
      TabIndex        =   7
      Top             =   900
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   300
      TabIndex        =   6
      Top             =   300
      Width           =   315
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   300
      Width           =   1515
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   900
      Width           =   1515
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   2100
      TabIndex        =   2
      Top             =   900
      Width           =   615
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   2100
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   0
      Left            =   2100
      TabIndex        =   0
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "frmHallOfFame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type Scores
    Name As String
    Scor As Integer
End Type
Dim HighScores(2) As Scores
Dim CheckUpdate As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim NewScore As Integer
    Dim Name As String

    If Not CheckUpdate Then
        NewScore = frmMain.NumberOfRows
    
CheckAgain:
        Open "highscores.dat" For Input As #1
            Do While Not EOF(1)
                Input #1, HighScores(i).Name
                Input #1, HighScores(i).Scor
                i = i + 1
            Loop
        Close #1
    
        
        For i = 0 To 2
            If Val(HighScores(i).Scor) <= NewScore Then
                Name = InputBox("Enter your name: ", "Power Tetris")
                Call UpdateHighScores(Name, i, NewScore)
                Call DisplayNames
                CheckUpdate = True
                Exit Sub
            End If
        Next i
End If
Call DisplayNames
 
Exit Sub
ErrorHandler:
    Close #1
    Call CreateNewFile
    GoTo CheckAgain
End Sub

Private Sub UpdateHighScores(Name As String, Position As Integer, NewScore As Integer)
    Dim i As Integer
    
    For i = 2 To Position + 1 Step -1
        HighScores(i).Name = HighScores(i - 1).Name
        HighScores(i).Scor = HighScores(i - 1).Scor
    Next i
    
    HighScores(Position).Name = Name
    HighScores(Position).Scor = NewScore
    
    Open "highscores.dat" For Output As #1
        For i = 0 To 2
            Print #1, HighScores(i).Name
            Print #1, HighScores(i).Scor
        Next i
    Close #1
End Sub

Private Sub CreateNewFile()
    Open "highscores.dat" For Output As #1
        Print #1, "N.N."
        Print #1, 10
        Print #1, "John Doe"
        Print #1, 5
        Print #1, "?"
        Print #1, 3
    Close #1
End Sub

Private Sub DisplayNames()
    Dim i As Integer
    
    For i = 0 To 2
        lblName(i) = HighScores(i).Name
        lblScore(i) = HighScores(i).Scor
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.DisplayHighScores = False
    CheckUpdate = False
End Sub
