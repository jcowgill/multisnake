VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmControl
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sock
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ipLookup As New Collection
Dim indexLookup As New Collection
Public nextIndex As Long

Private Sub Form_Load()
    indexLookup.Add "127.0.0.1", "0"
    ipLookup.Add 0, "127.0.0.1"

    sock.LocalPort = SNAKE_PORT
    sock.RemotePort = SNAKE_PORT
    sock.Bind

    nextIndex = 1
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
    'Read data
    Dim data() As Byte
    sock.GetData data

    'If not controller, pass on immediately
    If Not isController Then
        NetReceiveData 0, data
    Else
        'Lookup IP
        Dim index As Long
        On Error GoTo newUser
        index = ipLookup(sock.RemoteHostIP)
        On Error GoTo 0

        'Pass to dispatch function
        NetReceiveData index, data
        GoTo finish

newUser:
        'Unknown user sent a message
        If setupState = 0 And data(0) = MSG_CONNECT Then
            'Store in lookups
            indexLookup.Add sock.RemoteHostIP, CStr(nextIndex)
            ipLookup.Add nextIndex, sock.RemoteHostIP
            nextIndex = nextIndex + 1

            'Update counter
            clientsConnected = clientsConnected + 1
            UpdateCounter

            'Reply
            Send1Byte nextIndex - 1, MSG_CONNECTDONE
        End If

finish:
        sock.RemoteHost = "255.255.255.255"
    End If
End Sub

Public Sub RemoveSocket(ByVal index As Long)
    'Remove from tables
    ipLookup.Remove indexLookup(CStr(index))
    indexLookup.Remove index
End Sub

Public Sub SendSocket(ByVal index As Long, data() As Byte)
    If isController Then
        sock.RemoteHost = indexLookup(CStr(index))
        sock.SendData data
        sock.RemoteHost = "255.255.255.255"
    Else
        sock.SendData data
    End If
End Sub

Private Sub sock_Error(ByVal Number As Integer, Description As String, _
        ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, _
        ByVal HelpContext As Long, CancelDisplay As Boolean)

    'Ignore if not from main socket
    MsgBox "Socket error " & Number & vbCrLf & vbTab & Description, vbCritical, "MultiSnake"
    NetCloseAll
    End
End Sub
