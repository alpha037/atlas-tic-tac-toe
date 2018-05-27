VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic tac Toe - S.Dutta"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9525
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleMode       =   0  'User
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command12 
      BackColor       =   &H000000C0&
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Score Board"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3135
      Left            =   6120
      TabIndex        =   10
      Top             =   240
      Width           =   3255
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Make It Even"
         BeginProperty Font 
            Name            =   "Agency FB"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Player  O :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Player  X :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00808080&
      Caption         =   "&Start Over"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   3015
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3720
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3720
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Private Sub Command1_Click()
If flag = False Then
    Command1.Caption = "X"
    flag = True
    Label5.Caption = "Player O's Turn."
    Command1.Enabled = False
Else
    Command1.Caption = "O"
    flag = False
    Label5.Caption = "Player X's Turn."
    Command1.Enabled = False
End If
Win
End Sub

Private Sub Command10_Click()
Command1.Enabled = True
Command1.Caption = ""
Command2.Enabled = True
Command2.Caption = ""
Command3.Enabled = True
Command3.Caption = ""
Command4.Enabled = True
Command4.Caption = ""
Command5.Enabled = True
Command5.Caption = ""
Command6.Enabled = True
Command6.Caption = ""
Command7.Enabled = True
Command7.Caption = ""
Command8.Enabled = True
Command8.Caption = ""
Command9.Enabled = True
Command9.Caption = ""
End Sub

Private Sub Command11_Click()
Label3.Caption = ""
Label4.Caption = ""
Command1.Enabled = True
Command1.Caption = ""
Command2.Enabled = True
Command2.Caption = ""
Command3.Enabled = True
Command3.Caption = ""
Command4.Enabled = True
Command4.Caption = ""
Command5.Enabled = True
Command5.Caption = ""
Command6.Enabled = True
Command6.Caption = ""
Command7.Enabled = True
Command7.Caption = ""
Command8.Enabled = True
Command8.Caption = ""
Command9.Enabled = True
Command9.Caption = ""
End Sub

Private Sub Command12_Click()
MsgBox "SEE YOU LATER, COWARD !", vbCritical, "Quit Game"
End
End Sub

Private Sub Command2_Click()
If flag = False Then
    Command2.Caption = "X"
    flag = True
    Label5.Caption = "Player O's Turn."
    Command2.Enabled = False
Else
    Command2.Caption = "O"
    flag = False
    Label5.Caption = "Player X's Turn."
    Command2.Enabled = False
End If
Win
End Sub

Private Sub Command3_Click()
If flag = False Then
    Command3.Caption = "X"
    flag = True
    Label5.Caption = "Player O's Turn."
    Command3.Enabled = False
Else
    Command3.Caption = "O"
    flag = False
    Label5.Caption = "Player X's Turn."
    Command3.Enabled = False
End If
Win
End Sub

Private Sub Command4_Click()
If flag = False Then
    Command4.Caption = "X"
    flag = True
    Label5.Caption = "Player O's Turn."
    Command4.Enabled = False
Else
    Command4.Caption = "O"
    flag = False
    Label5.Caption = "Player X's Turn."
    Command4.Enabled = False
End If
Win
End Sub

Private Sub Command5_Click()
If flag = False Then
    Command5.Caption = "X"
    flag = True
    Label5.Caption = "Player O's Turn."
    Command5.Enabled = False
Else
    Command5.Caption = "O"
    flag = False
    Label5.Caption = "Player X's Turn."
    Command5.Enabled = False
End If
Win
End Sub

Private Sub Command6_Click()
If flag = False Then
    Command6.Caption = "X"
    flag = True
    Label5.Caption = "Player O's Turn."
    Command6.Enabled = False
Else
    Command6.Caption = "O"
    flag = False
    Label5.Caption = "Player X's Turn."
    Command6.Enabled = False
End If
Win
End Sub

Private Sub Command7_Click()
If flag = False Then
    Command7.Caption = "X"
    flag = True
    Label5.Caption = "Player O's Turn."
    Command7.Enabled = False
Else
    Command7.Caption = "O"
    flag = False
    Label5.Caption = "Player X's Turn."
    Command7.Enabled = False
End If
Win
End Sub

Private Sub Command8_Click()
If flag = False Then
    Command8.Caption = "X"
    flag = True
    Label5.Caption = "Player O's Turn."
    Command8.Enabled = False
Else
    Command8.Caption = "O"
    flag = False
    Label5.Caption = "Player X's Turn."
    Command8.Enabled = False
End If
Win
End Sub

Private Sub Command9_Click()
If flag = False Then
    Command9.Caption = "X"
    flag = True
    Label5.Caption = "Player O's Turn."
    Command9.Enabled = False
Else
    Command9.Caption = "O"
    flag = False
    Label5.Caption = "Player X's Turn."
    Command9.Enabled = False
End If
Win
End Sub

Private Sub Win()
If Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X" Then
MsgBox "WE HAVE A WINNER, AND THAT'S PLAYER X !", vbInformation, "Winner Winner, Chicken Dinner"
Label3.Caption = Val(Label3.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X" Then
MsgBox "WE HAVE A WINNER, AND THAT'S PLAYER X !", vbInformation, "Winner Winner, Chicken Dinner"
Label3.Caption = Val(Label3.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X" Then
MsgBox "WE HAVE A WINNER, AND THAT'S PLAYER X !", vbInformation, "Winner Winner, Chicken Dinner"
Label3.Caption = Val(Label3.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X" Then
MsgBox "WE HAVE A WINNER, AND THAT'S PLAYER X !", vbInformation, "Winner Winner, Chicken Dinner"
Label3.Caption = Val(Label3.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command2.Caption = "X" And Command5.Caption = "X" And Command8.Caption = "X" Then
MsgBox "WE HAVE A WINNER, AND THAT'S PLAYER X !", vbInformation, "Winner Winner, Chicken Dinner"

Label3.Caption = Val(Label3.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command3.Caption = "X" And Command6.Caption = "X" And Command9.Caption = "X" Then
MsgBox "WE HAVE A WINNER, AND THAT'S PLAYER X !", vbInformation, "Winner Winner, Chicken Dinner"
Label3.Caption = Val(Label3.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X" Then
MsgBox "WE HAVE A WINNER, AND THAT'S PLAYER X !", vbInformation, "Winner Winner, Chicken Dinner"
Label3.Caption = Val(Label3.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command3.Caption = "X" And Command5.Caption = "X" And Command7.Caption = "X" Then
MsgBox "WE HAVE A WINNER, AND THAT'S PLAYER X !", vbInformation, "Winner Winner, Chicken Dinner"
Label3.Caption = Val(Label3.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If




If Command1.Caption = "O" And Command2.Caption = "O" And Command3.Caption = "O" Then
MsgBox "WELL, IT LOOKS LIKE THE WINNER'S PLAYER O !", vbInformation, "Winner Winner, Chicken Dinner"
Label4.Caption = Val(Label4.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command4.Caption = "O" And Command5.Caption = "O" And Command6.Caption = "O" Then
MsgBox "WELL, IT LOOKS LIKE THE WINNER'S PLAYER O !", vbInformation, "Winner Winner, Chicken Dinner"
Label4.Caption = Val(Label4.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command7.Caption = "O" And Command8.Caption = "O" And Command9.Caption = "O" Then
MsgBox "WELL, IT LOOKS LIKE THE WINNER'S PLAYER O !", vbInformation, "Winner Winner, Chicken Dinner"
Label4.Caption = Val(Label4.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command1.Caption = "O" And Command4.Caption = "O" And Command7.Caption = "O" Then
MsgBox "WELL, IT LOOKS LIKE THE WINNER'S PLAYER O !", vbInformation, "Winner Winner, Chicken Dinner"
Label4.Caption = Val(Label4.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command2.Caption = "O" And Command5.Caption = "O" And Command8.Caption = "O" Then
MsgBox "WELL, IT LOOKS LIKE THE WINNER'S PLAYER O !", vbInformation, "Winner Winner, Chicken Dinner"
Label4.Caption = Val(Label4.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command3.Caption = "O" And Command6.Caption = "O" And Command9.Caption = "O" Then
MsgBox "WELL, IT LOOKS LIKE THE WINNER'S PLAYER O !", vbInformation, "Winner Winner, Chicken Dinner"
Label4.Caption = Val(Label4.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command1.Caption = "O" And Command5.Caption = "O" And Command9.Caption = "O" Then
MsgBox "WELL, IT LOOKS LIKE THE WINNER'S PLAYER O !", vbInformation, "Winner Winner, Chicken Dinner"
Label4.Caption = Val(Label4.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If

If Command3.Caption = "O" And Command5.Caption = "O" And Command7.Caption = "O" Then
MsgBox "WELL, IT LOOKS LIKE THE WINNER'S PLAYER O !", vbInformation, "Winner Winner, Chicken Dinner"
Label4.Caption = Val(Label4.Caption) + Val(1)
Command1.Caption = ""
Command2.Caption = ""
Command3.Caption = ""
Command4.Caption = ""
Command5.Caption = ""
Command6.Caption = ""
Command7.Caption = ""
Command8.Caption = ""
Command9.Caption = ""
End If



End Sub
