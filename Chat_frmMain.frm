VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "test程序"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "Chat_frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "清空"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3600
      Width           =   4935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1560
      Top             =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "发包"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Text            =   "192.168.5.116"
      Top             =   1680
      Width           =   2055
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Top             =   6465
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "没有连接"
            TextSave        =   "没有连接"
            Key             =   "STATUS"
            Object.Tag             =   ""
            Object.ToolTipText     =   "The current status of the connection"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6562
            TextSave        =   ""
            Key             =   "DATA"
            Object.Tag             =   ""
            Object.ToolTipText     =   "The last data transfer through the modem"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtRemotePort 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2010
      TabIndex        =   3
      Text            =   "4002"
      Top             =   720
      Width           =   1665
   End
   Begin VB.TextBox txtLocalPort 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2010
      TabIndex        =   2
      Text            =   "4001"
      Top             =   420
      Width           =   1665
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "连接"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   1365
   End
   Begin VB.TextBox txtRemoteIP 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2010
      TabIndex        =   1
      Top             =   120
      Width           =   1665
   End
   Begin VB.Frame Frame2 
      Caption         =   "本地IP"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   5025
   End
   Begin VB.Frame Frame1 
      Caption         =   "远程IP"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   5025
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   4080
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label4 
      Caption         =   "RemoteIP"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Remote Port :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   720
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Local Port :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   420
      Width           =   1905
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "本地IP："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   120
      Width           =   1905
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sendb(0 To 31) As Byte
Dim sendb1(0 To 31) As Byte

Dim aa(0 To 31) As String

Dim vta1
Dim vta2

Dim senddata(0 To 31, 0 To 2) As String
Dim Sendlong(0 To 31, 0 To 2) As Long
Dim IAPNO As Long

Dim CurrentIAP As Long
Dim mytemp(0 To 3) As Byte

Dim Rundirect As Integer





Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private IgnoreText As Boolean

Private Type MyType
      haha1 As Byte
      haha2 As Byte
      haha3 As Byte
      haha4 As Byte
      haha5 As Byte
      haha6 As Byte
      haha7 As Byte
      haha8 As Byte
      Mac1 As Byte
      Mac2 As Byte
      Mac3 As Byte
      Mac4 As Byte
      Mac5 As Byte
      Mac6 As Byte
      haha15 As Byte
      haha16 As Byte
      x1 As Byte
      x2 As Byte
      x3 As Byte
      x4 As Byte
      y1 As Byte
      y2 As Byte
      y3 As Byte
      y4 As Byte
      z1 As Byte
      z2 As Byte
      z3 As Byte
      z4 As Byte
      DT1 As Byte
      DT2 As Byte
      haha31 As Byte
      haha32 As Byte
End Type



Private Sub cmdClear_Click()

End Sub

Private Sub cmdConnect_Click()
On Error GoTo ErrHandler

With Winsock1
'  .RemoteHost = Trim(txtRemoteIP)
'  .RemotePort = Trim(txtRemotePort)

   If .LocalPort = Empty Then
      .LocalPort = Trim(txtLocalPort)
      Frame2.Caption = .LocalIP
      .Bind .LocalPort
      End If
End With

txtLocalPort.Locked = True

StatusBar1.Panels(1).Text = "  正在连接到 " & Winsock1.RemoteHost & "  "

Frame1.Enabled = True
Frame2.Enabled = True

StatusBar1.Panels(1).Text = "  连接成功 "
Exit Sub

ErrHandler:
    MsgBox "建立连接失败，按 F1 以获得帮助信息", vbCritical
End Sub



Private Sub Command1_Click()
Text1.Text = ""

End Sub

Private Sub Command2_Click()
 

Open App.Path + "\aaa.txt" For Input As #1

Dim j As Integer

For j = 0 To 31

Input #1, aa(j)

sendb(j) = CByte(aa(j))
sendb1(j) = CByte(aa(j))
'MsgBox aa(j)
Next


sendb1(12) = 144
sendb1(13) = 136


sendb(16) = 255
sendb(17) = 230
sendb(18) = 118
sendb(19) = 26

sendb1(16) = 255
sendb1(17) = 233
sendb1(18) = 233
sendb1(19) = 204


Winsock2.RemotePort = "1"
'Winsock2.LocalPort = "3101"



Close #1


Timer1.Enabled = True


End Sub





'当按下“F1”键时显示帮助信息
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
ChDir App.Path
'调用外部程序notepad.exe来打开帮助文本文件
Shell "notepad.exe readme.txt", vbNormalFocus
End If


End Sub

'当窗体加载时显示提示信息并在 txtRemoteIP 框中显示本地主机的IP
Private Sub Form_Load()
Show
txtRemoteIP = Winsock1.LocalIP


End Sub

Private Sub Timer1_Timer()



vta1 = sendb

Winsock2.RemoteHost = Text4.Text

Winsock2.RemotePort = txtRemotePort
Winsock2.senddata sendb


End Sub


'当 WINSOCK 接收到新的数据（信息）时，进行以下响应
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Open App.Path + "\writenew1.txt" For Append As #1

Dim vta
Dim mt As MyType
 Dim bt() As Byte
 
 Winsock1.GetData vta, vbArray + vbByte, LenB(mt)
 bt = vta

 
 Dim i As Integer
 For i = 0 To bytesTotal - 1
 Write #1, bt(i)
 Text1.Text = Text1.Text & bt(i) & " "
 Next
 Write #1, "next"
 


Frame1.Caption = Winsock1.RemoteHostIP
'在状态栏中显示接收信息
StatusBar1.Panels(2).Text = "  接收到 " & bytesTotal & " byte的消息  "

Close #1

End Sub


Private Function MyHex(aaa As Byte) As String

If Len(Hex(aaa)) = 1 Then
    MyHex = "0" & Hex(aaa)
Else
    MyHex = Hex(aaa)
End If

End Function

Private Function ComputeXYZ(a1 As Byte, a2 As Byte, a3 As Byte, a4 As Byte) As Long

Dim tempa As Long
Dim tempb As Long
Dim tempc As Long

If Len(Hex(a1)) = 2 And Left(Hex(a1), 1) = "F" Then

   'tempa = 4294967296# - tempa
    
End If

tempa = a1
tempa = tempa * 256
tempa = tempa * 256
tempa = tempa * 256

tempb = a2
tempb = tempb * 256
tempb = tempb * 256

tempc = a3
tempc = tempc * 256
tempc = tempc + a4

ComputeXYZ = tempa + tempb + tempc

End Function

Private Function HextoDec(tcHex As String) As String

tcHex = UCase(tcHex)

Dim n As Integer
Dim lnlen As Integer
Dim lnDec As Integer

Dim lcCurChr As String
Dim lnCurNum As Integer

lnlen = Len(tcHex)
lnDec = 0

For n = 1 To lnlen
lcCurChr = Mid(tcHex, n, 1)

If lcCurChr >= "A" Then
lnCurNum = Asc(lcCurChr) - 55
Else
lnCurNum = Val(lcCurChr)

End If
lnDec = lnCurNum * 16 ^ (lnlen - n) + lnDec

Next
HextoDec = Str(lnDec)

End Function

Private Sub metric_2_degree()
  
    Dim R, PI, xx, cos_b1 As Double
    Dim E_FLAT As Double
    Dim X, Y, Z As Double
    X = trackX
    Y = trackY
    Z = trackz
    
    'X = -2171242 lon=116.27
    'Y = 4398092  lan=39.83
    'Z = 4063680  alt=118.56
    On Error GoTo errhandle
    E_FLAT = (6.3781363 / 6.356742) * (6.3781363 / 6.356742)
    R = Sqr(X * X + Y * Y + Z * Z)
    PI = 4 * Atn(1)
    Dim rad As Double
    rad = 180 / PI
    xx = Sqr(X * X + Y * Y) / R
    cos_b1 = xx / Sqr((1 - E_FLAT * E_FLAT) * (xx * xx) + E_FLAT * E_FLAT)
    
    If (X <> 0) Then
       If (Y > 0 & X > 0) Or (Y < 0 & X > 0) Then
           trackX = rad * Atn(Y / X)
       Else
           trackX = rad * (Atn(Y / X) + PI)
           
       End If
    Else
        If Y > 1 Then
            trackX = rad * PI / 2
        Else
            trackX = -rad * PI / 2
        End If
    End If
    
    Dim Min As Double
    If (1 < cos_b1) Then
        Min = 1
    Else
        Min = cos_b1
    End If
    Dim Acos As Double
    Acos = Atn(-Min / Sqr(-Min * Min + 1)) + 2 * Atn(1)
    If (Z < 0) Then
        trackY = rad * Acos * (-1)
    Else
        trackY = rad * Acos * 1
    End If
    
    trackz = R - ((((0.003 * xx * xx + 0.7978) * xx * xx + 39.832) * xx * xx + 21353.6416) * xx * xx + 6356742.0252)
    Exit Sub
errhandle:
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End Sub

