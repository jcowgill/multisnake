Attribute VB_Name = "Network"
Option Explicit

Public Const SNAKE_PORT As Long = 50002

Public Const BLOCKS_HIGH As Long = 20

Public Enum ControllerMsgs
    'Messages sent by controller
    MSG_CONNECTDONE
    msg_disconnect
    MSG_ORDERSTART          'Signal that the order is being chosen
    MSG_ORDERRESPONSE       'Response to an order request <2 id>
    MSG_STARTGAME           'Snake game is starting <1 starting bit combination>
    MSG_SNAKEENTER          'Snake head has entered your computer <2 number of pieces (inc head), 1 position it entered>
    MSG_FOODEATEN           'The snake has eaten food <1 weather your should make new food>
    MSG_SNAKEDEAD           'The snake has run into itself and is dead (pause until disconnect or STARTGAME)
End Enum

Public Enum ClientMsgs
    'Messages sent by client
    MSG_CONNECT
    MSG_DISCONNECTSRV
    MSG_ORDERREQUEST        'Ask the server for an order number
    MSG_SNAKEEXIT           'Tell controller that the snake has exited <1 position>
    MSG_FOODEAT             'Tell controller that food on this computer has been eaten
    MSG_SIGNALDEAD          'Tell controller that the snake is dead
End Enum

Public Enum StartCombs
    'Starting bit combinations
    WITH_SNAKE = 1          'Starts with the snake
    WITH_FOOD = 2           'Starts with food
End Enum

'Added to vertical distance to calculate positions
Public Const LEFT_SIDE As Byte = 0
Public Const RIGHT_SIDE As Byte = 128


'Weather this is the controller
Public isController As Boolean

'Setup state
' 0 = Unconnected / Waiting for connections
' 1 = Connected (client only)
' 2 = Selecting order (client and server)
' 3 = Playing game
Public setupState As Integer

'Number of connected clients
Public clientsConnected As Integer

'Next order number (undefined if state <> 2)
Private nextOrderID As Integer

'Amount of food eaten
Public SnakePieceCount As Integer

'Conversions between IDs (only valid after ordering)
Private sockToPosition() As Integer
Private positionToSock() As Integer

'Starts ordering clients
Public Sub NetStartOrder()
    Dim i As Long

    'Wipe arrays
    ReDim sockToPosition(frmControl.nextIndex - 1)
    ReDim positionToSock(clientsConnected)

    'Start ordering
    nextOrderID = 1
    setupState = 2

    'Send message to everyone
    On Error Resume Next
    For i = 1 To UBound(sockToPosition)
        Send1Byte i, MSG_ORDERSTART
    Next

    Err.Clear
    On Error GoTo 0

    frmSetup.cmdGo.Enabled = False
    frmSetup.labStatus(0) = "Selecting order..."

End Sub

'Starts / restarts the game
Public Sub NetStartGame()
    'Reset numbers
    SnakePieceCount = 3
    setupState = 3

    'Select random computers for food and the snake start point
    Dim food, snake, i As Long
    Dim byteToSend As Byte

    food = Int((clientsConnected + 1) * Rnd)
    snake = Int((clientsConnected + 1) * Rnd)

    'Send data
    For i = 1 To clientsConnected
        byteToSend = 0

        If i = food Then
            byteToSend = byteToSend + WITH_FOOD
        End If

        If i = snake Then
            byteToSend = byteToSend + WITH_SNAKE
        End If

        Send2Bytes positionToSock(i), MSG_STARTGAME, byteToSend
    Next

    'Handle ourselves
    frmSnake.EventRestartGame (snake = 0), (food = 0)
End Sub

Public Sub NetCloseAll()
    'Disconnects everyone before dieing
    Dim i As Long

    On Error Resume Next

    If isController Then
        For i = 1 To frmControl.nextIndex - 1
            Send1Byte i, msg_disconnect
        Next
        Err.Clear
    Else
        Send1Byte 0, msg_disconnect
    End If

    On Error GoTo 0
End Sub

'Connect to controller
Public Sub NetConnect(controller As String)
    frmControl.sock.RemoteHost = controller
    Send1Byte 0, MSG_CONNECT
End Sub

'Event when data is received from a socket
Public Sub NetReceiveData(ByVal index As Integer, data() As Byte)
    Dim i As Long

    'Check there is a message
    On Error Resume Next
    Dim x As Integer
    x = UBound(data)

    If Err Then Exit Sub
    On Error GoTo 0

    'Process messages
    If isController Then
        Select Case data(0)
            Case MSG_DISCONNECTSRV
                If setupState = 0 Then
                    'Unload socket
                    frmControl.RemoveSocket index
                Else
                    'Die
                    NetCloseAll
                    End
                End If

            Case MSG_ORDERREQUEST        'Ask the server for an order number
                If setupState = 2 And sockToPosition(index) = 0 Then
                    'Generate order number and send it back
                    Send1ByteAndInt index, MSG_ORDERRESPONSE, nextOrderID
                    positionToSock(nextOrderID) = index
                    sockToPosition(index) = nextOrderID

                    'Check if this is the last person
                    If nextOrderID = clientsConnected Then
                        frmSetup.labStatus(0) = "Everyone has been ordered, press next to start the game"
                        frmSetup.cmdGo.Enabled = True
                    Else
                        nextOrderID = nextOrderID + 1
                    End If
                End If

            Case MSG_SNAKEEXIT           'Tell controller that the snake has exited <1 position>
                If setupState = 3 Then
                    'Raise event on controller
                    If data(1) >= RIGHT_SIDE Then
                        EventCSnakeExitControllerCode index, True, data(1) - RIGHT_SIDE
                    Else
                        EventCSnakeExitControllerCode index, False, data(1)
                    End If
                End If

            Case MSG_FOODEAT             'Tell controller that food on this computer has been eaten
                If setupState = 3 Then
                    'Raise event on controller
                    EventCFoodEat
                End If

            Case MSG_SIGNALDEAD          'Tell controller that the snake is dead
                If setupState = 3 Then
                    'Raise event on controller
                    EventCSnakeDead
                End If
        End Select
    Else
        Select Case data(0)
            Case MSG_CONNECTDONE
                'Display msg
                frmSetup.labStatus(0) = "Connected. Waiting for order to be chosen..."
                frmSetup.labID = "C"
                frmSetup.labID.ForeColor = vbGreen

                setupState = 1

            Case msg_disconnect
                'Die
                End

            Case MSG_ORDERSTART          'Signal that the order is being chosen
                'Ignore if not in connected state
                If setupState = 1 Then
                    'Wipe screen
                    setupState = 2
                    frmSetup.labStatus(0) = "Press the <space> key to select this computer's order"
                    frmSetup.labStatus(1) = "Order the computers CLOCKWISE"
                    frmSetup.labID = ""
                End If

            Case MSG_ORDERRESPONSE       'Response to an order request <2 id>
                'Ignore if not in choosing
                If setupState = 2 Then
                    'Wipe screen
                    setupState = 2
                    frmSetup.labStatus(0) = "Order selected"
                    frmSetup.labStatus(1) = ""
                    frmSetup.labID = ByteArrayToInteger(data, 1)
                End If

            Case MSG_STARTGAME           'Snake game is starting <1 starting bit combination>
                'Ignore if not choosing or playing
                If setupState = 2 Or setupState = 3 Then
                    frmSetup.Hide
                    frmSnake.Show

                    frmSnake.EventRestartGame (data(1) And WITH_SNAKE) <> 0, (data(1) And WITH_FOOD) <> 0

                    setupState = 3
                End If

            Case MSG_SNAKEENTER          'Snake head has entered your computer <2 number of pieces (inc head), 1 position it entered>
                If setupState = 3 Then
                    SnakePieceCount = ByteArrayToInteger(data, 1)

                    If data(3) >= RIGHT_SIDE Then
                        frmSnake.EventSnakeEnter True, data(3) - RIGHT_SIDE
                    Else
                        frmSnake.EventSnakeEnter False, data(3)
                    End If
                End If

            Case MSG_FOODEATEN           'The snake has eaten food (causes the tail to pause for 1 tick) <1 weather your should make new food>
                If setupState = 3 Then
                    SnakePieceCount = SnakePieceCount + 1
                    frmSnake.EventFoodEaten (data(1) = 1)
                End If

            Case MSG_SNAKEDEAD           'The snake has run into itself and is dead (pause until disconnect or STARTGAME)
                If setupState = 3 Then
                    frmSnake.EventSnakeDead
                End If

        End Select
    End If
End Sub

' ===============================
'  Game Events To Controller
' ===============================

Private Sub EventCSnakeExitControllerCode(ByVal index As Integer, ByVal onRight As Boolean, ByVal position As Integer)
    'Determine where to send the snake
    Dim nextPos As Long

    If onRight Then
        'Go right 1
        nextPos = sockToPosition(index) + 1
        If nextPos > clientsConnected Then nextPos = 0
    Else
        'Go left 1
        nextPos = sockToPosition(index) - 1
        If nextPos < 0 Then nextPos = clientsConnected
    End If

    'Send it to other computer
    If nextPos = 0 Then
        'Us!
        frmSnake.EventSnakeEnter Not onRight, position
    Else
        Dim outData(3) As Byte

        outData(0) = MSG_SNAKEENTER
        IntegerToByteArray outData, SnakePieceCount, 1

        If Not onRight Then
            outData(3) = position + RIGHT_SIDE
        Else
            outData(3) = position
        End If

        frmControl.SendSocket positionToSock(nextPos), outData
    End If
End Sub

'Tells the controller the  snake head has exited this computer
' See frmSnake.EventSnakeEnter for parameters
' Note: this may cause EventSnakeEnter to be called BEFORE this returns
Public Sub EventCSnakeExit(ByVal onRight As Boolean, ByVal position As Integer)
    If isController Then
        'Raise controller event from socket 0
        EventCSnakeExitControllerCode 0, onRight, position
    Else
        'Tell controller
        If onRight Then
            position = position + RIGHT_SIDE
        End If

        Send2Bytes 0, MSG_SNAKEEXIT, position
    End If
End Sub

'Tells the controller the food has been eaten on this computer
' This causes EventFoodEaten to be called immediately
Public Sub EventCFoodEat()
    'Controller?
    If isController Then
        'Pick a random computer to place the food
        Dim computer, i As Long
        computer = Int((clientsConnected + 1) * Rnd)

        'Send notifications to everyone
        For i = 1 To clientsConnected
            If i = computer Then
                Send2Bytes positionToSock(i), MSG_FOODEATEN, 1
            Else
                Send2Bytes positionToSock(i), MSG_FOODEATEN, 0
            End If
        Next

        'Update foood eaten count
        SnakePieceCount = SnakePieceCount + 1

        'Process ourselves
        frmSnake.EventFoodEaten (computer = 0)
    Else
        'Send to controller
        Send1Byte 0, MSG_FOODEAT
    End If
End Sub

'Tells the controller that the snake is dead
' Note: This calls EventSnakeDead immediately
Public Sub EventCSnakeDead()
    If isController Then
        'Broadcast to everyone
        Dim i As Long
        For i = 1 To clientsConnected
            Send1Byte positionToSock(i), MSG_SNAKEDEAD
        Next

        'Start death timer
        frmSetup.deathTimer.Enabled = True

        'Ourselves event
        frmSnake.EventSnakeDead
    Else
        Send1Byte 0, MSG_SIGNALDEAD
    End If
End Sub

'Update setup counter - clients connected
Public Sub UpdateCounter()
    Select Case clientsConnected
    Case 0
        frmSetup.labStatus(1) = "no clients connected"

    Case 1
        frmSetup.labStatus(1) = "1 client connected"

    Case Else
        frmSetup.labStatus(1) = clientsConnected & " clients connected"

    End Select
End Sub

Public Function ByteArrayToInteger(arr() As Byte, Optional off As Long = 0) As Integer
    'The network stuff uses big endian (most significant first)
    ByteArrayToInteger = arr(off) * 256 + arr(off + 1)
End Function

Public Sub IntegerToByteArray(arr() As Byte, ByVal data As Integer, Optional off As Long = 0)
    'The network stuff uses big endian (most significant first)
    arr(off) = data \ 256
    arr(off + 1) = data And &HFF
End Sub

Public Sub Send1Byte(ByVal index As Long, ByVal data1 As Byte)
    Dim dataArray(0) As Byte
    dataArray(0) = data1

    frmControl.SendSocket index, dataArray
End Sub

Public Sub Send2Bytes(ByVal index As Long, ByVal data1 As Byte, ByVal data2 As Byte)
    Dim dataArray(1) As Byte
    dataArray(0) = data1
    dataArray(1) = data2

    frmControl.SendSocket index, dataArray
End Sub

Public Sub Send1ByteAndInt(ByVal index As Long, ByVal data1 As Byte, ByVal data2 As Integer)
    Dim dataArray(2) As Byte
    dataArray(0) = data1
    IntegerToByteArray dataArray, data2, 1

    frmControl.SendSocket index, dataArray
End Sub
