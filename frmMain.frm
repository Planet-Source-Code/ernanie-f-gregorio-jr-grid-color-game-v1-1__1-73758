VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grid Number Game"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
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
   ScaleHeight     =   4980
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timGameProgress 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3420
      Top             =   4020
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      Height          =   480
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   450
   End
   Begin VB.CommandButton cmdBackTo 
      Caption         =   "Back to Main menu"
      Height          =   480
      Left            =   1650
      TabIndex        =   2
      Top             =   4020
      Width           =   1650
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   4020
      Width           =   1650
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 Minute"
      Height          =   195
      Left            =   540
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   750
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iLevel    As Integer
Public iRotation As Integer
Dim iHour        As Integer
Dim iMinute      As Integer
Dim iSecond      As Integer
Dim lColr(5)     As Long
Dim blnShuffling As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
'Private Sub LockLabels(ByVal prmEnable As Boolean)
'    Dim intCount As Integer
'
'    For intCount = 0 To (iLevel * iLevel) - 1
'        lblMain(intCount).Enabled = prmEnable
'    Next intCount
'    timGameProgress.Enabled = prmEnable
'End Sub

Private Function PuzzleFinished() As Boolean
    Dim intCount As Long
    PuzzleFinished = True

    For intCount = lblMain.LBound To lblMain.UBound
        If lblMain(intCount).Caption <> "1" Then
            PuzzleFinished = False
            Exit For
        End If
    Next intCount
End Function

Private Sub RandomizeGrid()
    Dim intCount     As Integer
    Dim intColorCode As Integer
    lblTime.Caption = "Shuffling..."
    timGameProgress.Enabled = False

    For intCount = 0 To (iLevel * iLevel) - 1 ' reset the grid to red
        lblMain(intCount).Caption = 1
        lblMain(intCount).BackColor = lColr(lblMain(intCount).Caption)
        lblMain(intCount).ForeColor = lColr(lblMain(intCount).Caption)
    Next intCount
    blnShuffling = True

    For intCount = 1 To 10 ' shuffle the grid by clicking the label

        DoEvents
        Randomize
        Call lblMain_Click(Rnd * ((iLevel ^ 2) - 1))
        Sleep 100 ' give the player a clue on what is clicked although it take sharp eyes and good memory to memorize the random move that was generated. :)
    Next intCount
    blnShuffling = False
    ' reset time
    iHour = 0
    iMinute = 0
    iSecond = 0
    lblTime.Caption = "Start!!!"
    timGameProgress.Enabled = True ' timer start! :)
End Sub

Private Sub RotateValue(ByVal intIndex As Long)
    lblMain(intIndex).Caption = (lblMain(intIndex).Caption Mod iRotation) + 1
    lblMain(intIndex).BackColor = lColr(lblMain(intIndex).Caption)
    lblMain(intIndex).ForeColor = lColr(lblMain(intIndex).Caption)
End Sub

Private Sub cmdCheck_Click()
    If PuzzleFinished Then
        MsgBox "Congratulations!! You finished the puzzle."
    End If
End Sub

Private Sub cmdBackTo_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    MsgBox "The aim of the game is to make the grid a solid red color."
End Sub

Private Sub cmdReset_Click()
    Call RandomizeGrid
End Sub

Private Sub Form_Load()
    ' width   height
    GenerateGrid lblMain, iLevel, iLevel, 0
    lColr(1) = vbRed
    lColr(2) = vbBlue
    lColr(3) = vbYellow
    lColr(4) = vbGreen
    lColr(5) = vbWhite
    '    Call RandomizeGrid ' will shuffle the grid
    Me.Width = (lblMain(0).Width) * iLevel ' set the width of the form by multiplying the width of the label and the number of cols
    Me.Height = (lblMain(0).Height * iLevel) + cmdReset.Height + cmdHelp.Height + 400
    ' set the size of the buttons
    cmdReset.Width = Me.ScaleWidth / 2
    cmdBackTo.Width = Me.ScaleWidth / 2
    ' set the position of the buttons
    cmdBackTo.Left = cmdReset.Width
    cmdReset.Top = lblMain(0).Top + (lblMain(0).Height * iLevel)
    cmdBackTo.Top = lblMain(0).Top + (lblMain(0).Height * iLevel)
    Call RandomizeGrid
End Sub

Private Sub lblMain_Click(Index As Integer)
    'If timGameProgress.Enabled = True Then
    If Index >= iLevel Then
        RotateValue Index - iLevel ' above
    End If
    If Index Mod iLevel Then
        RotateValue Index - 1 ' left
    End If
    If Index Mod iLevel < iLevel - 1 Then
        RotateValue Index + 1 'right
    End If
    If Index < (iLevel * (iLevel - 1)) Then
        RotateValue Index + iLevel 'below
    End If
    RotateValue Index
    If PuzzleFinished And Not blnShuffling Then
        MsgBox "Congratulations!! You finished the puzzle in " & lblTime.Caption
        lblTime.Caption = "Congratulations!! "
        timGameProgress.Enabled = False
        MsgBox "Click reset to start a new game with the same settings." & vbCrLf & "Click Back to main menu to select different level."
    End If
    'End If
End Sub

Private Sub timGameProgress_Timer()
    If iSecond < 59 Then
        iSecond = iSecond + 1
    Else
        iSecond = 0
        If iMinute < 59 Then
            iMinute = iMinute + 1
        Else
            iMinute = 0
            iHour = iHour + 1
        End If
    End If
    lblTime.Caption = iHour & " Hours " & iMinute & " minute/s " & iSecond & " second/s"
End Sub
