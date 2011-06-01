VERSION 5.00
Begin VB.Form frmSnake
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Snake"
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer timGameLoop
      Enabled         =   0   'False
      Left            =   2880
      Top             =   2040
   End
   Begin VB.Shape shpFood
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   2160
      Shape           =   1  'Square
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape shpBlock
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   1080
      Shape           =   1  'Square
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmSnake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INTERVAL As Long = 100

Dim lastKeyPress As Integer
Dim nextKeyPress As Integer

Dim blockSize As Long
Dim blocksWide As Long

Dim headIsOnScreen As Long

Dim tailBlock As Long
Dim headBlock As Long

'The tags of all blocks contain the LIFETIME of the block in clock ticks
' Each timer pulse decreases all block lifetimes by 1. Anything with 0 life left is destroyed
'

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If it it a up, down, left, right - use it
    Select Case KeyCode
    Case vbKeyUp, vbKeyW
        If lastKeyPress <> vbKeyDown Then
            nextKeyPress = vbKeyUp
        End If

    Case vbKeyDown, vbKeyS

        If lastKeyPress <> vbKeyUp Then
            nextKeyPress = vbKeyDown
        End If

    Case vbKeyRight, vbKeyD
        If lastKeyPress <> vbKeyLeft Then
            nextKeyPress = vbKeyRight
        End If

    Case vbKeyLeft, vbKeyA
        If lastKeyPress <> vbKeyRight Then
            nextKeyPress = vbKeyLeft
        End If

    Case vbKeyEscape
        'Quit
        End

    End Select
End Sub

Private Sub MoveFoodOnThisComputer()
    'Move food to random location
    Dim blockX As Long
    Dim blockY As Long
    Dim i As Long
    Dim continue As Boolean

    shpFood.Visible = False

    'Pick random numbers in loop
    Do
        'Pick a random block from 0 to SCREEN_<h/w> - 1
        blockX = Int((blocksWide - 1) * Rnd)
        blockY = Int((BLOCKS_HIGH - 1) * Rnd)

        'Move the food
        shpFood.Left = blockSize * blockX + ((blockSize - shpFood.Width) / 2)
        shpFood.Top = blockSize * blockY + ((blockSize - shpFood.Height) / 2)

        'Check if it's in the snake
        continue = False
        For i = tailBlock To headBlock
            If CheckCollision(shpBlock(i), shpFood) Then
                continue = True
                Exit For
            End If
        Next
    Loop While continue

    shpFood.Visible = True

End Sub

Private Function CheckCollision(ctrl1 As Control, ctrl2 As Control) As Boolean
    CheckCollision = ctrl1.Left + ctrl1.Width > ctrl2.Left And _
                     ctrl2.Left + ctrl2.Width > ctrl1.Left And _
                     ctrl1.Top + ctrl1.Height > ctrl2.Top And _
                     ctrl2.Top + ctrl2.Height > ctrl1.Top
End Function

Private Sub Form_Load()
    Randomize
End Sub

Private Sub timGameLoop_Timer()
    'Do block life draining
    Dim i As Long

    For i = tailBlock To headBlock
        shpBlock(i).Tag = Int(shpBlock(i).Tag - 1)

        If shpBlock(i).Tag = 0 Then
            'Kill
            Unload shpBlock(i)
            tailBlock = tailBlock + 1
        End If
    Next

    'Process head block
    If headIsOnScreen Then
        'Set tail colour
        shpBlock(headBlock).FillColor = vbYellow

        'Determine new position for the head
        Dim newX As Long    'In blocks
        Dim newY As Long    'In blocks
        newX = shpBlock(headBlock).Left / blockSize
        newY = shpBlock(headBlock).Top / blockSize

        Select Case nextKeyPress
        Case vbKeyUp
            newY = newY - 1

            'Check if off screen
            If newY < 0 Then
                'Jump to bottom
                newY = BLOCKS_HIGH - 1
            End If

        Case vbKeyDown
            newY = newY + 1

            'Check if off screen
            If newY >= BLOCKS_HIGH Then
                'Jump to top
                newY = 0
            End If

        Case vbKeyLeft
            newX = newX - 1

            'Check if off screen
            If newX < 0 Then
                'Notify snake exit, then exit sub
                headIsOnScreen = False
                EventCSnakeExit False, newY
                Exit Sub
            End If

        Case vbKeyRight
            newX = newX + 1

            'Check if off screen
            If newX >= blocksWide Then
                'Notify snake exit, then exit sub
                headIsOnScreen = False
                EventCSnakeExit True, newY
                Exit Sub
            End If
        End Select

        'Load new head block / notify of snake exit
        MakeNewHead newX, newY

        'Check for collision with ourselves
        For i = tailBlock To headBlock - 1
            If CheckCollision(shpBlock(headBlock), shpBlock(i)) Then
                EventCSnakeDead
                Exit Sub
            End If
        Next

        'Check for food collision
        If shpFood.Visible And CheckCollision(shpBlock(headBlock), shpFood) Then
            'Eat the food
            shpFood.Visible = False
            EventCFoodEat
        End If
    End If

    'Copy lastKeyPress
    lastKeyPress = nextKeyPress
End Sub

'Creates a new head block at the given block position
Private Sub MakeNewHead(ByVal blockX As Integer, ByVal blockY As Integer)
    headBlock = headBlock + 1
    Load shpBlock(headBlock)

    With shpBlock(headBlock)
        'Set block lifetime
        .Tag = SnakePieceCount

        'Colour and position
        .FillColor = vbBlue
        .Move blockX * blockSize, blockY * blockSize
        .Visible = True
    End With
End Sub

' ======================================
'  Game Events
' ======================================

'Occurs after food has been eaten on any computer
' makeFood tells weather food should be created on this computer
Public Sub EventFoodEaten(ByVal makeFood As Boolean)
    'Pause snake for 1 tick
    Dim i As Long
    For i = tailBlock To headBlock
        shpBlock(i).Tag = Int(shpBlock(i).Tag + 1)
    Next

    'Make food
    If makeFood Then MoveFoodOnThisComputer
End Sub

'Occurs when the snake's head enters this computer
' isRight is true when the snake enteres from the right
' position is the block from the stop the snake entered at (top block is 0)
' Number of pieces in the snake is from SnakePieceCount
Public Sub EventSnakeEnter(ByVal isRight As Boolean, ByVal position As Integer)
    'Get X coordinate + update current keys
    Dim blockX As Integer
    If isRight Then
        lastKeyPress = vbKeyLeft
        nextKeyPress = vbKeyLeft
        blockX = blocksWide - 1
    Else
        lastKeyPress = vbKeyRight
        nextKeyPress = vbKeyRight
        blockX = 0
    End If

    'Make new head
    headIsOnScreen = True
    MakeNewHead blockX, position

End Sub

'Occurs when the snake is dead
Public Sub EventSnakeDead()
    'If we have the snake, colour head red
    If headIsOnScreen Then
        shpBlock(headBlock).FillColor = vbRed
        shpBlock(headBlock).ZOrder
    End If

    'Stop timer
    timGameLoop.Enabled = False
End Sub

'Occurs when the game is started and restarted
Public Sub EventRestartGame(ByVal makeSnake As Boolean, ByVal makeFood As Boolean)
    'Calculate block size
    blockSize = ScaleHeight / BLOCKS_HIGH
    blocksWide = ScaleWidth / blockSize

    'Hide food
    shpFood.Visible = False
    shpFood.Width = blockSize / 2
    shpFood.Height = blockSize / 2

    'Set block template size
    shpBlock(0).Height = blockSize
    shpBlock(0).Width = blockSize

    'Reset stuff
    lastKeyPress = vbKeyUp
    nextKeyPress = vbKeyUp

    'Show snake screen and hide the snake and the food
    Dim i As Long
    If headBlock > 0 Then
        For i = tailBlock To headBlock
            Unload shpBlock(i)
        Next
    End If
    tailBlock = 1
    headBlock = 0

    shpFood.Visible = False

    'Create snake if nessesary
    If makeSnake Then
        'Put snake in middle
        Load shpBlock(1)
        Load shpBlock(2)
        Load shpBlock(3)

        shpBlock(2).Top = Int((BLOCKS_HIGH - 1) / 2) * blockSize
        shpBlock(2).Left = Int((blocksWide - 1) / 2) * blockSize
        shpBlock(2).Visible = True
        shpBlock(3).FillColor = vbYellow
        shpBlock(2).Tag = 2

        shpBlock(1).Left = shpBlock(2).Left
        shpBlock(1).Top = shpBlock(2).Top + blockSize
        shpBlock(1).Visible = True
        shpBlock(3).FillColor = vbYellow
        shpBlock(1).Tag = 1

        shpBlock(3).Left = shpBlock(2).Left
        shpBlock(3).Top = shpBlock(2).Top - blockSize
        shpBlock(3).Visible = True
        shpBlock(3).FillColor = vbBlue
        shpBlock(3).Tag = 3

        headBlock = 3

        headIsOnScreen = True
    Else
        headIsOnScreen = False
    End If

    'Create food if nessesary
    If makeFood Then
        MoveFoodOnThisComputer
    End If

    'Restart timer
    timGameLoop.INTERVAL = 0
    timGameLoop.INTERVAL = INTERVAL
    timGameLoop.Enabled = True
    Show
End Sub
