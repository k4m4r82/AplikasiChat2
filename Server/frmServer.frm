VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "Aplikasi Server"
   ClientHeight    =   4005
   ClientLeft      =   3180
   ClientTop       =   2085
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " [ Komputer Laen ] "
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3855
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.TextBox txtChat 
      Height          =   285
      Left            =   4650
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   4650
      TabIndex        =   0
      Top             =   3450
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   4080
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMain 
      Height          =   2685
      Left            =   4650
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   645
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Pesan"
      Height          =   195
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Width           =   450
   End
End
Attribute VB_Name = "frmServer"
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

Private Sub Winsock1_Close(Index As Integer)
    List1.List(Index - 1) = Winsock1(Index).RemoteHostIP & " on port " & Winsock1(Index).RemotePort & " [disconnected]"
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i           As Long
    Dim j           As Long
    Dim newWinsock  As Boolean
    
    On Error GoTo errHandle
        
    If Index = 0 Then 'ingat yang bertugas untuk mengecek permintaan koneksi adalah objek winsock dengan index = 0
        'ini bagian yang bertugas untuk mengecek winsock yang idle
        For i = 1 To Winsock1.UBound
            If (Winsock1(i).State = sckClosed) Or (Winsock1(i).State = sckClosing) Then
                j = i   'var j menampung index winsock yang idle
                Exit For
            End If
        Next i
      
        'jika j = 0 berarti semua winsock kepakai, otomatis kita harus menambahkan winsock yang baru
        If j = 0 Then
            Call Load(Winsock1(Winsock1.UBound + 1))
            j = Winsock1.UBound
            newWinsock = True
        End If
        
        'terima koneksi yang baru
        With Winsock1(j)
            Call .Close
            Call .Accept(requestID)
        End With
        
        If newWinsock Then
            List1.AddItem Winsock1(j).RemoteHostIP & " on port " & Winsock1(j).RemotePort & " [connected]"
        Else
            List1.List(j - 1) = Winsock1(j).RemoteHostIP & " on port " & Winsock1(j).RemotePort & " [connected]"
        End If
        
    End If
    
    Exit Sub
    
errHandle:
    Call Winsock1(0).Close
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String

    'get the data and display it in the textbox
    Winsock1(Index).GetData strData
    txtMain.Text = txtMain.Text & vbCrLf & Winsock1(Index).RemoteHostIP & " say : " & strData
    txtMain.SelStart = Len(txtMain.Text)
End Sub

Private Sub cmdSend_Click()
    If List1.ListIndex < 0 Then
        MsgBox "Komputer laen belum dipilih", vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    Winsock1(List1.ListIndex + 1).SendData txtChat.Text
    DoEvents
    
    txtMain.Text = txtMain.Text & vbCrLf & "U say : " & txtChat.Text
    txtChat.Text = ""
End Sub

Private Sub Form_Load()
    With Winsock1(0)
        .Close
        .LocalPort = LOCAL_PORT
        .Listen
    End With
End Sub

