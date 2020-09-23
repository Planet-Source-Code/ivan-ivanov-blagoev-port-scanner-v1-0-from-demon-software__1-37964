VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demon Software - Port Scanner v1.0"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton Command1 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton ConnectBTN 
         Caption         =   "Scan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton StopScan 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox RemoteIP_1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "127"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox RemoteIP_2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox RemoteIP_3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox PortStart 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox PortEnd 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "65535"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox RemoteIP_4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TimeOutInterval 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Text            =   "500"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "IP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Ports:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Time Out:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4200
      Top             =   0
   End
   Begin VB.TextBox Status 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2760
      Width           =   5775
   End
   Begin MSWinsockLib.Winsock myTCPclient 
      Left            =   4800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Line Line4 
      X1              =   5520
      X2              =   5520
      Y1              =   2040
      Y2              =   2400
   End
   Begin VB.Line Line3 
      X1              =   5520
      X2              =   2760
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      X1              =   2760
      X2              =   2760
      Y1              =   2400
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   5520
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label CurrntPort 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Demon Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Current port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Result:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Port1 As Long
Dim Port2 As Long
Dim IPnum As String
Dim ScanBroi As Long
Dim ScannStr As Boolean

Private Sub Command1_Click()
MsgBox ("Created by Ivan Blagoev - 'Demon Software'.")
End Sub

Private Sub Form_Load()
On Error GoTo Err_Hand

ScanBroi = 0
ScannStr = False
Exit Sub

Err_Hand:
MsgBox Err.Description
End Sub

Private Sub StopScan_Click()
Rem Cancel Scanning
If ScannStr = True Then
Me.TimeOut.Enabled = False
myTCPclient.Close
MsgBox ("Abort scanning.")
Me.ConnectBTN.Enabled = True
Me.RemoteIP_1.Enabled = True
Me.RemoteIP_2.Enabled = True
Me.RemoteIP_3.Enabled = True
Me.RemoteIP_4.Enabled = True
Me.TimeOutInterval.Enabled = True
Me.PortStart.Enabled = True
Me.PortEnd.Enabled = True
ScannStr = False
End If
End Sub

Rem Scann Button
Private Sub ConnectBTN_Click()
On Error GoTo Err_Hand

Rem Verify values IP and Ports
On Error GoTo Err_IP
If Me.RemoteIP_1.Text < 0 Or Me.RemoteIP_1.Text > 255 Or Me.RemoteIP_1.Text = "" Then GoTo Err_IP
If Me.RemoteIP_2.Text < 0 Or Me.RemoteIP_2.Text > 255 Or Me.RemoteIP_2.Text = "" Then GoTo Err_IP
If Me.RemoteIP_3.Text < 0 Or Me.RemoteIP_3.Text > 255 Or Me.RemoteIP_3.Text = "" Then GoTo Err_IP
If Me.RemoteIP_4.Text < 0 Or Me.RemoteIP_4.Text > 255 Or Me.RemoteIP_4.Text = "" Then GoTo Err_IP

On Error GoTo No_Port
If Val(Me.PortStart.Text) < 1 Or Val(Me.PortStart.Text) > 65535 Then GoTo No_Port
If Val(Me.PortEnd.Text) < 1 Or Val(Me.PortEnd.Text) > 65535 Then GoTo No_Port
If Val(Me.PortStart.Text) > Val(Me.PortEnd.Text) Then GoTo No_Port

On Error GoTo Err_TimeOut
If Me.TimeOutInterval.Text < 10 Or Me.TimeOutInterval.Text > 10000 Or Me.TimeOutInterval.Text = "" Then GoTo Err_TimeOut

On Error GoTo Err_Hand
Rem Value entry
Port1 = Me.PortStart.Text
Port2 = Me.PortEnd.Text
IPnum = Me.RemoteIP_1.Text & "." & Me.RemoteIP_2.Text & "." _
& Me.RemoteIP_3.Text & "." & Me.RemoteIP_4.Text

Rem Change Menu and run scanning
Me.Status.Text = ""
Me.CurrntPort.Caption = ""
ScanBroi = Port1
Me.StopScan.SetFocus
Me.ConnectBTN.Enabled = False
Me.RemoteIP_1.Enabled = False
Me.RemoteIP_2.Enabled = False
Me.RemoteIP_3.Enabled = False
Me.RemoteIP_4.Enabled = False
Me.TimeOutInterval.Enabled = False
Me.PortStart.Enabled = False
Me.PortEnd.Enabled = False
ScannStr = True
Me.TimeOut.Interval = Me.TimeOutInterval.Text
Call SannProc

Exit Sub

Rem Errors in procedure
No_Port:
MsgBox ("No valid port entry.")
Exit Sub

Err_IP:
MsgBox ("No valid IP address.")
Exit Sub

Err_TimeOut:
MsgBox ("Time out interval is: 10 to 10000")
Exit Sub

Err_Hand:
MsgBox Err.Description & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
Rem Close connection
Me.myTCPclient.Close
End Sub

Private Sub ShowStatus()
On Error GoTo Err_Hand
Dim Str2 As String
Dim Stat As Byte

Stat = myTCPclient.State

Rem Connection Status
Select Case Stat
Case 0
            Exit Sub
Case 1
            Me.Status.Text = Me.Status.Text & vbCrLf & ScanBroi - 1 & " - Open" & vbCrLf
Case 2
            Me.Status.Text = Me.Status.Text & vbCrLf & ScanBroi - 1 & " - Listening" & vbCrLf
Case 3
            Me.Status.Text = Me.Status.Text & vbCrLf & ScanBroi - 1 & " - Connection pending" & vbCrLf
Case 4
            Me.Status.Text = Me.Status.Text & vbCrLf & ScanBroi - 1 & " -  Resolving host" & vbCrLf
Case 5
            Me.Status.Text = Me.Status.Text & vbCrLf & ScanBroi - 1 & " -  Host resolved" & vbCrLf
Case 6
            Exit Sub
Case 7
            Me.myTCPclient.GetData Str2, , 150
            If Str2 = "" Then
            Me.Status.Text = Me.Status.Text & vbCrLf & ScanBroi - 1 & " -  Connected" & vbCrLf
            Exit Sub
            End If
            Me.Status.Text = Me.Status.Text & vbCrLf & ScanBroi - 1 & " -  Connected: " & Str2 & vbCrLf
Case 8
            Me.Status.Text = Me.Status.Text & vbCrLf & ScanBroi - 1 & " -  Peer is closing the connection" & vbCrLf
Case 9
            Me.Status.Text = Me.Status.Text & vbCrLf & ScanBroi - 1 & " -  Error" & vbCrLf
End Select

Exit Sub

Err_Hand:
MsgBox Err.Description
End Sub

Private Sub Err_IP()
Rem IP message
On Error GoTo Err_Hand
MsgBox ("No valid IP: 0 - 255")
Exit Sub

Err_Hand:
MsgBox Err.Description
End Sub

Private Sub TimeOut_Timer()
On Error GoTo Err_Hand
Me.TimeOut.Enabled = False

Call ShowStatus
Call SannProc

Rem Port + 1 ***** Port counter value
ScanBroi = ScanBroi + 1

Rem If all ports scanning then EXIT
If ScanBroi > Port2 Then
Me.TimeOut.Enabled = False
myTCPclient.Close
MsgBox ("End scanning.")
Me.ConnectBTN.Enabled = True
Me.RemoteIP_1.Enabled = True
Me.RemoteIP_2.Enabled = True
Me.RemoteIP_3.Enabled = True
Me.RemoteIP_4.Enabled = True
Me.TimeOutInterval.Enabled = True
Me.PortStart.Enabled = True
Me.PortEnd.Enabled = True
ScannStr = False
Exit Sub
End If
Exit Sub

Err_Hand:
MsgBox Err.Description
End Sub

Private Sub SannProc()
On Error GoTo Err_Hand

Rem Close Connection
If (myTCPclient.State <> sckClosed) Then myTCPclient.Close

Rem View current port
Me.CurrntPort.Caption = ScanBroi

Rem Connect IP and Port
Me.myTCPclient.RemoteHost = IPnum
Me.myTCPclient.RemotePort = ScanBroi
Me.myTCPclient.Connect

Rem Enabling timer for Scanning
Me.TimeOut.Enabled = True
Exit Sub

Err_Hand:
MsgBox Err.Description
End Sub
