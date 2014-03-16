VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmKlien 
   Caption         =   "Aplikasi Klien"
   ClientHeight    =   3105
   ClientLeft      =   9975
   ClientTop       =   2085
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   690
      TabIndex        =   5
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   375
      Left            =   690
      TabIndex        =   3
      Top             =   2610
      Width           =   975
   End
   Begin VB.TextBox txtChat 
      Height          =   285
      Left            =   690
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4080
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMain 
      Height          =   1485
      Left            =   690
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1005
      Width           =   5655
   End
   Begin VB.Label Label2 
      Caption         =   "Server"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Pesan"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   450
   End
End
Attribute VB_Name = "frmKlien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Private Const LOCAL_PORT As Long = 11111

Private Function startConnect(ByVal ipServer As String) As Boolean
    Dim i   As Long
    
    'On Error Resume Next
    
    If Winsock1.State <> sckClosed Then Winsock1.Close ' close existing connection
    Call Winsock1.Connect(ipServer, LOCAL_PORT)
    With Winsock1
        Do While .State <> sckConnected
            DoEvents
            If .State = sckError Then Exit Function
        Loop
    End With
    
    startConnect = True
End Function

Private Sub cmdConnect_Click()
    txtMain.Text = ""
    txtChat.Text = ""
    
    If cmdConnect.Caption = "Connect" Then
        If startConnect(txtServer.Text) Then
            cmdConnect.Caption = "Disconnet"
            cmdSend.Enabled = True
        Else
            cmdSend.Enabled = False
        End If
    Else
        cmdSend.Enabled = False
        Winsock1.Close
        cmdConnect.Caption = "Connect"
    End If
End Sub

Private Sub cmdSend_Click()
    'send the data thats in the text box and
    'clear it to prepare for the next chat message
    Winsock1.SendData txtChat.Text
    DoEvents
    
    txtMain.Text = txtMain.Text & vbCrLf & "U say : " & txtChat.Text
    txtChat.Text = ""
End Sub

Private Sub Winsock1_Connect()
    'we are connected!
    'MsgBox "Connected"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    ' get the data from the socket
    Winsock1.GetData strData
    ' display it in the textbox
    txtMain.Text = txtMain.Text & vbCrLf & "Server say : " & strData
    ' scroll the box down
    txtMain.SelStart = Len(txtMain.Text)
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' an error has occured somewhere, so let the user know
    MsgBox "Error: " & Description
    ' close the socket, ready to go again
    Winsock1.Close
End Sub

