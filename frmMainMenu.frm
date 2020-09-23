VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grid Game v1.1"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox comRotation 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmMainMenu.frx":0000
      Left            =   3135
      List            =   "frmMainMenu.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start!!"
      Height          =   600
      Left            =   75
      TabIndex        =   3
      Top             =   1800
      Width           =   4410
   End
   Begin VB.ComboBox comCols 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmMainMenu.frx":0004
      Left            =   3135
      List            =   "frmMainMenu.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   780
      Width           =   1215
   End
   Begin VB.Label lblNumberOf 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numbers of color rotation >>>>"
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   375
      TabIndex        =   5
      Top             =   1260
      Width           =   2355
   End
   Begin VB.Label lblSelectDifficulty 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select level >>>"
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   1515
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblTheAim 
      BackStyle       =   0  'Transparent
      Caption         =   "The aim of the game is to make the grid a solid red color."
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   195
      TabIndex        =   1
      Top             =   300
      Width           =   4185
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
    frmMain.iLevel = comCols.Text
    frmMain.iRotation = comRotation.Text
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Dim iCount As Integer

    For iCount = 3 To 5
        comCols.AddItem iCount
    Next iCount

    For iCount = 2 To 5
        comRotation.AddItem iCount
    Next iCount
    ' default level
    comCols.Text = 4
    comRotation.Text = 2
End Sub
