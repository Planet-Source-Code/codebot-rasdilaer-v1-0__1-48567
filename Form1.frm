VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   435
      Left            =   645
      TabIndex        =   3
      Top             =   3705
      Width           =   1755
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1740
      TabIndex        =   2
      Top             =   3000
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   165
      TabIndex        =   1
      Top             =   3000
      Width           =   1170
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   2955
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, ByVal pSrc As String, ByVal ByteLen As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Const RAS95_MaxEntryName = 256
Const RAS_MaxPhoneNumber = 128
Const RAS_MaxCallbackNumber = RAS_MaxPhoneNumber

Const UNLEN = 256
Const PWLEN = 256
Const DNLEN = 12
Private Type RASDIALPARAMS
   dwSize As Long ' 1052
   szEntryName(RAS95_MaxEntryName) As Byte
   szPhoneNumber(RAS_MaxPhoneNumber) As Byte
   szCallbackNumber(RAS_MaxCallbackNumber) As Byte
   szUserName(UNLEN) As Byte
   szPassword(PWLEN) As Byte
   szDomain(DNLEN) As Byte
End Type

Private Type RASENTRYNAME95
    'set dwsize to 264
    dwSize As Long
    szEntryName(RAS95_MaxEntryName) As Byte
End Type

Private Declare Function RasDial Lib "rasapi32.dll" Alias "RasDialA" (ByVal lprasdialextensions As Long, ByVal lpcstr As String, ByRef lprasdialparamsa As RASDIALPARAMS, ByVal dword As Long, lpvoid As Any, ByRef lphrasconn As Long) As Long
Private Declare Function RasEnumEntries Lib "rasapi32.dll" Alias "RasEnumEntriesA" (ByVal reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long
Private Declare Function RasGetEntryDialParams Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" (ByVal lpcstr As String, ByRef lprasdialparamsa As RASDIALPARAMS, ByRef lpbool As Long) As Long

Private Function Dial(ByVal Connection As String, ByVal UserName As String, ByVal Password As String) As Boolean
    Dim rp As RASDIALPARAMS, h As Long, resp As Long
    rp.dwSize = Len(rp) + 6
    ChangeBytes Connection, rp.szEntryName
    ChangeBytes "", rp.szPhoneNumber 'Phone number stored for the connection
    ChangeBytes "*", rp.szCallbackNumber 'Callback number stored for the connection
    ChangeBytes UserName, rp.szUserName
    ChangeBytes Password, rp.szPassword
    ChangeBytes "*", rp.szDomain 'Domain stored for the connection
    'Dial
    resp = RasDial(ByVal 0, ByVal 0, rp, 0, ByVal 0, h)   'AddressOf RasDialFunc
    Dial = (resp = 0)
End Function

Private Function ChangeToStringUni(Bytes() As Byte) As String
    'Changes an byte array  to a Visual Basic unicode string
    Dim temp As String
    temp = StrConv(Bytes, vbUnicode)
    ChangeToStringUni = Left(temp, InStr(temp, Chr(0)) - 1)
End Function

Private Function ChangeBytes(ByVal str As String, Bytes() As Byte) As Boolean
    'Changes a Visual Basic unicode string to an byte array
    'Returns True if it truncates str
    Dim lenBs As Long 'length of the byte array
    Dim lenStr As Long 'length of the string
    lenBs = UBound(Bytes) - LBound(Bytes)
    lenStr = LenB(StrConv(str, vbFromUnicode))
    If lenBs > lenStr Then
        CopyMemory Bytes(0), str, lenStr
        ZeroMemory Bytes(lenStr), lenBs - lenStr
    ElseIf lenBs = lenStr Then
        CopyMemory Bytes(0), str, lenStr
    Else
        CopyMemory Bytes(0), str, lenBs 'Queda truncado
        ChangeBytes = True
    End If
End Function

Private Sub Command1_Click()
    Dial List1.Text, Text1, Text2
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Movable Form
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub List1_Click()
    Dim rdp As RASDIALPARAMS, t As Long
    rdp.dwSize = Len(rdp) + 6
    ChangeBytes List1.Text, rdp.szEntryName
    'Get User name and password for the connection
    t = RasGetEntryDialParams(List1.Text, rdp, 0)
    If t = 0 Then
        Text1 = ChangeToStringUni(rdp.szUserName)
        Text2 = ChangeToStringUni(rdp.szPassword)
    End If
End Sub

Private Sub Form_Load()
    'load the connections
    Text2.PasswordChar = "*"
    Command1.Caption = "Dial"
    Dim s As Long, l As Long, ln As Long, a$
    ReDim r(255) As RASENTRYNAME95
    
    r(0).dwSize = 264
    s = 256 * r(0).dwSize
    l = RasEnumEntries(vbNullString, vbNullString, r(0), s, ln)
    For l = 0 To ln - 1
        a$ = StrConv(r(l).szEntryName(), vbUnicode)
        List1.AddItem Left$(a$, InStr(a$, Chr$(0)) - 1)
    Next
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
        List1_Click
    End If
End Sub

