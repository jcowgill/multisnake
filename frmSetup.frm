VERSION 5.00
Begin VB.Form frmSetup
   BackColor       =   &H00000000&
   Caption         =   "Multi Snake"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9255
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer deathTimer
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7800
      Top             =   1320
   End
   Begin VB.CommandButton cmdGo
      Caption         =   "Next"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label labID
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font
         Name            =   "Arial"
         Size            =   399.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7395
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   10440
   End
   Begin VB.Label labStatus
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1035
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   10440
   End
   Begin VB.Label labStatus
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1995
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10440
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()
    'Select operation
    Select Case setupState
        Case 0
            'Move to selecting order
            If clientsConnected > 0 Then
                NetStartOrder
            End If

        Case 2
            'Start game
            Hide
            NetStartGame

    End Select
End Sub

Private Sub deathTimer_Timer()
    'Disable timer and restart
    deathTimer.Enabled = False
    NetStartGame
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            NetCloseAll

        Case vbKeySpace
            'If a drone and selecting order, select it
            If Not isController And setupState = 2 Then
                Send1Byte 0, MSG_ORDERREQUEST
            End If

    End Select
End Sub

Private Sub Form_Load()
    Show

    'Host or controller
    If MsgBox("Are you the controlling computer?", vbYesNo, "Multi Snake") = vbYes Then
        'Setup controller
        isController = True
        Load frmControl

        'Display IP
        labStatus(0) = "You are the controller and your name / ip is " & frmControl.sock.LocalIP
        labStatus(1) = "no computers are connected"
        cmdGo.Visible = True

    Else
        isController = False

        'Connect to controller
        Dim controller As String
        controller = InputBox("What is the controllers name / ip address?", "Multi Snake")

        If Len(controller) = 0 Then End

        labStatus(0) = "Connecting..."
        labID = "X"

        'Load control drone
        Load frmControl
        NetConnect controller
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NetCloseAll
    End
End Sub
