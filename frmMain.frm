VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP PortForward"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wskListen 
      Left            =   5265
      Top             =   1035
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstActive 
      Height          =   2985
      Left            =   225
      TabIndex        =   7
      Top             =   1845
      Width           =   6945
   End
   Begin VB.CommandButton btStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   5985
      TabIndex        =   6
      Top             =   90
      Width           =   1140
   End
   Begin VB.TextBox txtRemoteHost 
      Height          =   300
      Left            =   2205
      TabIndex        =   5
      Top             =   720
      Width           =   1725
   End
   Begin VB.TextBox txtRemotePort 
      Height          =   300
      Left            =   2205
      TabIndex        =   4
      Top             =   1035
      Width           =   645
   End
   Begin VB.TextBox txtDestPort 
      Height          =   300
      Left            =   2205
      TabIndex        =   3
      Top             =   135
      Width           =   645
   End
   Begin MSWinsockLib.Winsock wskLocalConn 
      Index           =   0
      Left            =   5940
      Top             =   1035
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskRemoteConn 
      Index           =   0
      Left            =   5940
      Top             =   585
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Active connections:"
      Height          =   240
      Index           =   3
      Left            =   225
      TabIndex        =   8
      Top             =   1620
      Width           =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote port:"
      Height          =   240
      Index           =   2
      Left            =   225
      TabIndex        =   2
      Top             =   1035
      Width           =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote host:"
      Height          =   240
      Index           =   1
      Left            =   225
      TabIndex        =   1
      Top             =   720
      Width           =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Local port:"
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   2040
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Long

Private Sub Form_Load()
    If Command <> "" Then
        Dim cmd() As String
        cmd = Split(Command, " ")
        txtDestPort.Text = cmd(0)
        txtRemoteHost.Text = cmd(1)
        txtRemotePort.Text = cmd(2)
        Call btStart_Click
    End If
End Sub


Private Sub btStart_Click()
    wskListen.LocalPort = txtDestPort.Text
    Call wskListen.Listen
    btStart.Enabled = False
    txtDestPort.Enabled = False
    txtRemoteHost.Enabled = False
    txtRemotePort.Enabled = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call wskListen.Close
End Sub


Private Sub wskListen_ConnectionRequest(ByVal requestID As Long)
    Dim newLocalConn As Winsock
    Dim newRemoteConn As Winsock
    Dim c As Long
    Dim index As Integer
    counter = counter + 1
    
    Call Load(wskRemoteConn(counter))
    Set newRemoteConn = wskRemoteConn(counter)
    newRemoteConn.RemoteHost = txtRemoteHost.Text
    newRemoteConn.RemotePort = txtRemotePort.Text
    
    Call newRemoteConn.Connect

    While Not newRemoteConn.State = MSWinsockLib.sckConnected And c < 1000
        DoEvents
        c = c + 1
    Wend
    
    If newRemoteConn.State = MSWinsockLib.sckConnected Then
        Call Load(wskLocalConn(counter))
        Set newLocalConn = wskLocalConn(counter)
        Call lstActive.AddItem(counter & vbTab & wskListen.RemoteHostIP & vbTab & Now())
        Call newLocalConn.Accept(requestID)
    Else
        Call newRemoteConn.Close
        Call Unload(wskRemoteConn)
        
    End If
    
End Sub


Private Sub wskLocalConn_Close(index As Integer)
    Call deleteItem(index)
    Call wskRemoteConn(index).Close
    Call Unload(wskLocalConn(index))
    Call Unload(wskRemoteConn(index))
    
End Sub


Private Sub wskRemoteConn_Close(index As Integer)
    Call deleteItem(index)
    Call wskLocalConn(index).Close
    Call Unload(wskLocalConn(index))
    Call Unload(wskRemoteConn(index))
    
End Sub


Private Sub wskRemoteConn_DataArrival(index As Integer, ByVal bytesTotal As Long)
    Dim sData As Variant
    Call wskRemoteConn(index).GetData(sData)
    Call wskLocalConn(index).SendData(sData)
    Call updateItem(index)
    
End Sub


Private Sub wskLocalConn_DataArrival(index As Integer, ByVal bytesTotal As Long)
    Dim sData As Variant
    Call wskLocalConn(index).GetData(sData)
    Call wskRemoteConn(index).SendData(sData)
    Call updateItem(index)
    
End Sub


Sub deleteItem(index)
    For i = 0 To lstActive.ListCount - 1
        Dim item() As String
        item = Split(lstActive.List(i), vbTab)
        If index = item(0) Then
            Call lstActive.RemoveItem(i)
        End If
    Next
    
End Sub


Sub updateItem(index)
    For i = 0 To lstActive.ListCount - 1
        Dim item() As String
        item = Split(lstActive.List(i), vbTab)
        If index = item(0) Then
            lstActive.List(i) = item(0) & vbTab & item(1) & vbTab & item(2) & vbTab & Now()
        End If
    Next
    
End Sub

