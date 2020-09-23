Attribute VB_Name = "modGridArray"
Option Explicit
Public Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long

Public Sub GenerateGrid(ctrl As Variant, _
                        ByVal GridW As Long, _
                        ByVal GridH As Long, _
                        Optional ByVal Sep As Long = 0)
    'Assumptions
    '1 ctrl is a single control on the form with an Index of 0
    '2 ctrl(0) is at the top left corner of the area you wish to cover with the grid
    '3 Ctrl(0) is correct size
    '
    'Sep allows you to separate the controls by a specified ammount
    '
    'Draws a Left2Right Top2Bottom grid
    'if you want your grid to flow differently fiddle with the RowPos/ColPos variables
    '
    Dim I      As Long
    Dim GridL  As Long  'Left of the Grid used to reset location when grid cycles
    Dim TileW  As Long  'used to increment the Left and Top positions of any tile
    Dim TileH  As Long  '(if tile is not square then they will not be identical)
    Dim RowPos As Long  'top of row of tiles
    Dim ColPos As Long  'left of current col of tiles

    On Error Resume Next

    'Set initial values for variables(calling the control in the loop is slower)
    LockWindowUpdate ctrl(0).Parent.hWnd

    With ctrl(0)
        TileW = .Width + Sep
        TileH = .Height + Sep
        RowPos = .Top - TileW 'trick to allow the Mod code to work on first row
        GridL = .Left
    End With

    'clean up
    For I = 0 To ctrl.UBound
        Unload ctrl(I)
        ' this removes the old array, if no array then it does nothing
    Next                            ' (NOTE 1 To.. you can't unload the seed control) I I I I I

    'position the controls
    For I = 0 To (GridW * GridH - 1) ' -1 because controls are zero-based
        If I > 0 Then                  ' you can't Load the initial control
            Load ctrl(I)
        End If
        ColPos = ColPos + TileW        'increase col by 1
        If I Mod GridW = 0 Then
            RowPos = RowPos + TileH      ' increase row by 1
            ColPos = GridL               ' reset col to 1st
        End If
        ctrl(I).Move ColPos, RowPos
        ctrl(I).Visible = True       'All 'Load'ed controls are Visible = False, so force them to show
    Next I
    LockWindowUpdate 0

    On Error GoTo 0

End Sub
':)Code Fixer V4.0.29 (Wednesday, 04 January 2006 11:20:27) 1 + 55 = 56 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 133302322223333232|33332222222222222222222222222202|1112222|2221222|222222222233|111111111111|1122222222220|333333|
