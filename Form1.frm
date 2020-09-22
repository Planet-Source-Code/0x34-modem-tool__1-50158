VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "0x34's Modem Coding Tool - For the MultiTECH DSVD V/Modem"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton NUMBER 
      Caption         =   "Dial #5"
      Height          =   375
      Left            =   1680
      TabIndex        =   51
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   4320
      Top             =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "VOICE REGISTERS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   36
      Top             =   2520
      Width           =   3015
      Begin VB.CommandButton VRX 
         BackColor       =   &H00FFC0C0&
         Caption         =   "VRX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton VTX 
         BackColor       =   &H00C0E0FF&
         Caption         =   "VTX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton LastSTAT 
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton VBS 
         Caption         =   "VBS ?"
         Height          =   375
         Left            =   1560
         TabIndex        =   47
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton VCI 
         Caption         =   "VCI ?"
         Height          =   375
         Left            =   1560
         TabIndex        =   46
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton VLS 
         Caption         =   "VLS ?"
         Height          =   375
         Left            =   1560
         TabIndex        =   45
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton VGT 
         Caption         =   "VGT ?"
         Height          =   375
         Left            =   1560
         TabIndex        =   44
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton VTD 
         Caption         =   "VTD ?"
         Height          =   375
         Left            =   360
         TabIndex        =   43
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton VSS 
         Caption         =   "VSS ?"
         Height          =   375
         Left            =   360
         TabIndex        =   42
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton VSR 
         Caption         =   "VSR ?"
         Height          =   375
         Left            =   360
         TabIndex        =   41
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton VSP 
         Caption         =   "VSP ?"
         Height          =   375
         Left            =   360
         TabIndex        =   40
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton CID 
         Caption         =   "CID ?"
         Height          =   375
         Left            =   360
         TabIndex        =   39
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton VRN 
         Caption         =   "VRN ?"
         Height          =   375
         Left            =   360
         TabIndex        =   38
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton VRA 
         Caption         =   "VRA ?"
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Factory"
      Height          =   315
      Left            =   6960
      TabIndex        =   35
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send this String"
      Height          =   315
      Left            =   5040
      TabIndex        =   34
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load MultiTECH Default"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7920
      TabIndex        =   33
      Top             =   8520
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   0
      TabIndex        =   31
      Text            =   "Text3"
      Top             =   8880
      Width           =   10815
   End
   Begin VB.CommandButton MSC 
      Caption         =   "Dial #4"
      Height          =   375
      Left            =   1680
      TabIndex        =   27
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton CLS8 
      Caption         =   "#CLS=8"
      Height          =   315
      Left            =   3960
      TabIndex        =   26
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton CellPhone 
      Caption         =   "Dial #3"
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Top             =   840
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   3360
      Top             =   960
   End
   Begin VB.CommandButton DialHOME 
      Caption         =   "Dial#2"
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton SelfTest 
      Caption         =   "Dial #1"
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton DisConnect 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame Registers 
      Caption         =   "MSComm1 Status"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   6720
      Width           =   3015
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   2580
         TabIndex        =   30
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   2580
         TabIndex        =   29
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   2580
         TabIndex        =   28
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label8 
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Stripped ="
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "ReSponse #3 ="
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "ReSponse #2 ="
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "ReSponse #1 ="
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Result 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   1500
         Width           =   2775
      End
      Begin VB.Label Result 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1365
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Result 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1365
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Result 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1365
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   410
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "MSComm Hits this cycle = "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.CommandButton ExeComm 
      Caption         =   "Execute Command"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6240
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Text"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton GetS86 
      BackColor       =   &H8000000A&
      Caption         =   "S86"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   900
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   8295
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   7575
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3480
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   3
      RThreshold      =   2
      RTSEnable       =   -1  'True
      BaudRate        =   56000
      SThreshold      =   1
   End
   Begin VB.CommandButton ConnectModem 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Connected"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton ATH 
      Caption         =   "ATH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton ATDT 
      Caption         =   "ATDT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton AT 
      Caption         =   "AT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Initialization String:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   8640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "AT Command to execute"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -120
      TabIndex        =   8
      Top             =   5640
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'         MODEM Communication & Protocol Development Tool
'                         Coded by 0x34
'
'          Open Source Code -  For anyone who can use it

Dim AccCntr As Integer
Dim Command As String
Dim S As Integer
Dim ModStr As String
Dim Stripped As Integer
Dim TEMP As String
Dim STAT As Boolean

Private Sub DEFAULT()
Text3 = "ATE0&K6Q0X4V0S0=0S30=0-SMS=0-SSE=0#CID=2#CLS=8#VRN=0#VLS=6#VBS=8"
End Sub

Private Sub AT_Click()
CCNTR
Command = "AT"
ExecuteStr
End Sub

Private Sub ATDT_Click()
CCNTR
Command = "ATDT"
ExecuteStr
End Sub

Private Sub ATH_Click()
CCNTR
Command = "+++ATH"
ExecuteStr
End Sub

Private Sub CellPhone_Click()
CCNTR
Command = "ATDT" ' Insert Test number here
ExecuteStr
End Sub

Private Sub CID_Click()
CCNTR
Command = "AT#CID?"
ExecuteStr
End Sub

Private Sub CLS8_Click()
CCNTR
Command = "AT#CLS=8#VRN=0#VLS=6"
ExecuteStr
End Sub

Private Sub Command1_Click()
Text1 = ""
CCNTR
Label3 = AccCntr
Result(0) = ""
Result(1) = ""
Result(2) = ""
Result(3) = ""
Label9(0) = ""
Label9(1) = ""
Label9(2) = ""
End Sub

Private Sub Command2_Click()
DEFAULT
End Sub

Private Sub Command3_Click()
CCNTR
Command = Text3
ExecuteStr
End Sub

Private Sub Command4_Click()
CCNTR
Command = "AT&F"
ExecuteStr
Text3 = "AT&F"
End Sub

Private Sub ConnectModem_Click()
CCNTR
StartUP
ConnectModem.BackColor = &HC0FFC0
ConnectModem.Caption = "Connected"
End Sub

Private Sub DialHOME_Click()
CCNTR
Command = "ATDT" ' Insert a test number here
ExecuteStr
End Sub

Private Sub DisConnect_Click()
CCNTR
ConnectModem.BackColor = &HC0C0FF
ConnectModem.Caption = "Connect Modem"
MSComm1.PortOpen = False
Text1 = Text1 + "------------------------" & vbNewLine
Text1 = Text1 + "MSComm Port Closed" & vbNewLine
Label3 = AccCntr
DisConnect.Enabled = False
Result(3) = "DISCONNECTED"
End Sub

Private Sub ExeComm_Click()
If Text2 = "" Then Exit Sub
CCNTR
Command = "AT" & Text2
ExecuteStr
End Sub

Private Sub Form_Load()
Timer2.Enabled = False
Timer1.Enabled = False
STAT = False
Text2 = ""
S = 0
Stripped = 0
DEFAULT
CCNTR
StartUP
Label8 = Stripped
End Sub

Private Sub StartUP()   'Connect the MODEM
Dim CONF As String
'Text1 = ""
CONF = "56000,N,8,1"
If MSComm1.PortOpen = True Then GoTo OPENPORT
MSComm1.CommPort = 1
MSComm1.Settings = CONF
MSComm1.PortOpen = True
'The port is now open
OPENPORT:
DisConnect.Enabled = True
Command = Text3
'Command = "ATV1Q0&A0S14.1=0"            'USRobotics
ExecuteStr
LabQ = "Q = 0"
LabX = "X = 4"
End Sub

Private Sub GetS86_Click()
CCNTR
Command = "ATS86?"
ExecuteStr
End Sub

Private Sub LastSTAT_Click()
CCNTR
STAT = True
Command = "AT&V1"
ExecuteStr
Timer2.Enabled = True
LastSTAT.BackColor = vbRed
End Sub

Private Sub MSC_Click()
CCNTR
Command = "ATDT" ' Insert a test number here
ExecuteStr
End Sub

Private Sub MSComm1_OnComm()
TEMP = ""
ModStr = ""
ModStr = CStr(MSComm1.Input)
TEMP = ModStr
If STAT Then
    Text1 = Text1 + ModStr
    AccCntr = AccCntr + 1
    Label3 = AccCntr
    Exit Sub
End If
Cleaner
If AccCntr < 3 Then
    Result(AccCntr) = TEMP
    Label9(AccCntr) = Stripped
End If
If ModStr = "" Then
    If Stripped > 0 Then
        'Text1 = Text1 + "------------------------" & vbNewLine
        Text1 = Text1 + "Response # " & (AccCntr + 1) & " = Control Characters (" & Stripped & ") Only" & vbNewLine
        AccCntr = AccCntr + 1
        Label3 = AccCntr
        Exit Sub
    Else
        'Text1 = Text1 + "------------------------" & vbNewLine
        Text1 = Text1 + "Response # " & (AccCntr + 1) & " = Empty Return" & vbNewLine
        AccCntr = AccCntr + 1
        Label3 = AccCntr
        Exit Sub
    End If
End If
'Text1 = Text1 + "------------------------" & vbNewLine
Text1 = Text1 + "Response # " & (AccCntr + 1) & " = " & ModStr
    If Stripped > 0 Then
        Text1 = Text1 + "   (Stripped " & Stripped & ")" & vbNewLine
    Else
        Text1 = Text1 + " " & vbNewLine
    End If
On Error GoTo ERROR
Result(3) = ""
If ModStr = "3" Then Result(3) = "NO CARRIER": FlashS86
If ModStr = "0" Then Result(3) = "OK"
If ModStr = "0200" Then Result(3) = "Key Abort Disconnect"
If ModStr = "0020" Then Result(3) = "NORMAL DISCONNECT"
If ModStr = "0210" Then Result(3) = "Clr Prev Dis Reason"
If ModStr = "7" Then Result(3) = "BUSY"
If ModStr = "6" Then Result(3) = "NO DIALTONE"
If ModStr = "1" Then Result(3) = "CONNECT"
If ModStr = "2" Then Result(3) = "RING"
If ModStr = "4" Then Result(3) = "ERROR"
If ModStr = "32" Then Result(3) = "NUMBER BLACKLISTED"
If ModStr = "33" Then Result(3) = "FAX MODE"
If ModStr = "35" Then Result(3) = "DATA CONNECTION"
If ModStr = "8" Then Result(3) = "NO ANSWER"
If ModStr = "VCON" Then Result(3) = "VOICE CONNECTION"
If ModStr = "0040" Then Result(3) = "PHYSICAL LOSS OF CARRIER"
AccCntr = AccCntr + 1
Label3 = AccCntr
If Result(3) = "" Then
    Result(3) = "UTL Result Code"
End If
Exit Sub
ERROR:
    Text1 = Text1 + "CODE DETERMINATION ERROR OCCURRED"
AccCntr = AccCntr + 1
Label3 = AccCntr
End Sub

Private Sub NUMBER_Click()
CCNTR
Command = "ATDT" ' Insert a Test number here
ExecuteStr
End Sub

Private Sub SelfTest_Click()
CCNTR
Command = "ATDT" ' Insert a Test number here
ExecuteStr
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ExeComm_Click
        Text2 = ""
    End If
End Sub
Private Sub CCNTR()
AccCntr = 0
Label9(0) = ""
Label9(1) = ""
Label9(2) = ""
End Sub
Private Sub ExecuteStr()
On Error GoTo ERR
Result(0) = ""
Result(1) = ""
Result(2) = ""
Result(3) = ""
Text1 = Text1 + "------------------------" & vbNewLine
Text1 = Text1 + "Send String = " & Command & vbNewLine
MSComm1.Output = Command & vbCr
Exit Sub
ERR:
    Text1 = Text1 + "------------------------" & vbNewLine
    Text1 = Text1 + "Cannot Execute Command (" & Command & ")" & vbNewLine
    Text1 = Text1 + "MSComm1.PortOpen = FALSE" & vbNewLine
End Sub
Private Sub FlashS86()
Timer1.Enabled = True
S = 0
GetS86.BackColor = vbRed
End Sub

Private Sub Timer1_Timer()
If GetS86.BackColor = vbRed Then
    GetS86.BackColor = &H8000000F
    S = S + 1
Else
    GetS86.BackColor = vbRed
    S = S + 1
End If
If S > 8 Then
    GetS86.BackColor = &H8000000F
    Timer1.Enabled = False
End If
End Sub

Private Sub Cleaner()
If ModStr = "" Then Exit Sub
Dim q As Integer
Dim P As Integer
Dim m As String
Dim K As String
Dim TMPmodStr As String
Stripped = 0
TMPmodStr = ""
q = Len(ModStr)
For i = 0 To q - 1
    K = Left$(ModStr, i)
    If K = "" Then GoTo Nutn
    m = Right$(K, 1)
    P = Asc(m)
    If P > 32 Then
        TMPmodStr = TMPmodStr + m
    Else
        Stripped = Stripped + 1
    End If
Nutn:
Next
ModStr = TMPmodStr
Label8 = Stripped
End Sub

Private Sub Timer2_Timer()
STAT = False
Timer2.Enabled = False
LastSTAT.BackColor = &H8000000F
End Sub

Private Sub VBS_Click()
CCNTR
Command = "AT#VBS?"
ExecuteStr
End Sub

Private Sub VCI_Click()
CCNTR
Command = "AT#VCI?"
ExecuteStr
End Sub

Private Sub VGT_Click()
CCNTR
Command = "AT#VGT?"
ExecuteStr
End Sub

Private Sub VLS_Click()
CCNTR
Command = "AT#VLS?"
ExecuteStr
End Sub

Private Sub VRA_Click()
CCNTR
Command = "AT#VRA?"
ExecuteStr
End Sub

Private Sub VRN_Click()
CCNTR
Command = "AT#VRN?"
ExecuteStr
End Sub

Private Sub VRX_Click()
CCNTR
Command = "AT#VRX"
ExecuteStr
End Sub

Private Sub VSP_Click()
CCNTR
Command = "AT#VSP?"
ExecuteStr
End Sub

Private Sub VSR_Click()
CCNTR
Command = "AT#VSR?"
ExecuteStr
End Sub

Private Sub VSS_Click()
CCNTR
Command = "AT#VSS?"
ExecuteStr
End Sub

Private Sub VTD_Click()
CCNTR
Command = "AT#VTD?"
ExecuteStr
End Sub

Private Sub VTX_Click()
CCNTR
Command = "AT#VTX"
ExecuteStr
End Sub
