VERSION 5.00
Begin VB.Form FrmDrag 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drag And Drop Program "
   ClientHeight    =   5850
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   5190
   ControlBox      =   0   'False
   Icon            =   "FrmDrag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      DragIcon        =   "FrmDrag.frx":000C
      Height          =   855
      Index           =   5
      Left            =   3480
      Picture         =   "FrmDrag.frx":044E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      DragIcon        =   "FrmDrag.frx":0890
      Height          =   855
      Index           =   4
      Left            =   120
      Picture         =   "FrmDrag.frx":0CD2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      DragIcon        =   "FrmDrag.frx":1114
      Height          =   855
      Index           =   3
      Left            =   120
      Picture         =   "FrmDrag.frx":1556
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      DragIcon        =   "FrmDrag.frx":1998
      Height          =   855
      Index           =   2
      Left            =   1800
      Picture         =   "FrmDrag.frx":1DDA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      DragIcon        =   "FrmDrag.frx":221C
      Height          =   855
      Index           =   1
      Left            =   3480
      Picture         =   "FrmDrag.frx":265E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      DragIcon        =   "FrmDrag.frx":2AA0
      Height          =   855
      Index           =   0
      Left            =   1800
      Picture         =   "FrmDrag.frx":2EE2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   855
      Index           =   5
      Left            =   3480
      TabIndex        =   17
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   855
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   855
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   855
      Index           =   2
      Left            =   1800
      TabIndex        =   14
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   855
      Index           =   1
      Left            =   3480
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   855
      Index           =   0
      Left            =   1800
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Japan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Index           =   5
      Left            =   3480
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Italy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Germany"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Canada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "France"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "America"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Menu MnuStart 
      Caption         =   "Start"
   End
   Begin VB.Menu MnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "FrmDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Dim l As String
Dim d As Integer
Dim counter As Integer
Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1(Index).Drag vbBeginDrag 'to begin the drag procedure
a = Index 'initializing a to index
End Sub

Private Sub Form_DragOver(source As Control, X As Single, Y As Single, State As Integer)
source.Move Label2(a).Left, Label2(a).Top
End Sub

Private Sub Form_Load()
Label1(0).Visible = False
Label2(0).Visible = False
Command1(0).Visible = False
Label1(1).Visible = False
Label2(1).Visible = False
Command1(1).Visible = False
Label1(2).Visible = False
Label2(2).Visible = False
Command1(2).Visible = False
Label1(3).Visible = False
Label2(3).Visible = False
Command1(3).Visible = False
Label1(4).Visible = False
Label2(4).Visible = False
Command1(4).Visible = False
Label1(5).Visible = False
Label2(5).Visible = False
Command1(5).Visible = False
End Sub

Private Sub Label1_DragOver(Index As Integer, Flag As Control, X As Single, Y As Single, State As Integer)
Flag.Move Label1(Index).Left, Label1(Index).Top
d = Index
If a <> d Then
Flag.Move Label2(a).Left, Label2(a).Top
End If
 If a = d Then
    counter = counter + 1
    End If
    If counter = 12 Then
    l = MsgBox("CONGRATULATIONS!!", vbOKOnly, "VERY GOOD")
    Beep
    counter = 1
    End If
         
End Sub

Private Sub Label2_DragOver(Index As Integer, source As Control, X As Single, Y As Single, State As Integer)
source.Move Label2(Index).Left, Label2(Index).Top
End Sub

Private Sub MnuExit_Click()
Call Unload(FrmDrag)
End Sub

Private Sub MnuStart_Click()
Label1(0).Visible = True
Label2(0).Visible = True
Command1(0).Visible = True
Label1(1).Visible = True
Label2(1).Visible = True
Command1(1).Visible = True
Label1(2).Visible = True
Label2(2).Visible = True
Command1(2).Visible = True
Label1(3).Visible = True
Label2(3).Visible = True
Command1(3).Visible = True
Label1(4).Visible = True
Label2(4).Visible = True
Command1(4).Visible = True
Label1(5).Visible = True
Label2(5).Visible = True
Command1(5).Visible = True
End Sub
