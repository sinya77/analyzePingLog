VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// ツール -> オプション -> 全般 のエラートラップを
'// 「クラス モジュールで中断」

Option Explicit

'// プロパティ
Public logTime As String
Public logIpMask As String
Public logRes As String

Private m_time As Date
Private m_ip As String
Private m_mask As Integer
Private m_nwAddr As String
Private m_res As Integer


' 初期化処理
Private Sub Class_Initialize()
  m_time = 0
  m_ip = ""
  m_mask = 0
  m_nwAddr = ""
  m_res = 0
End Sub

'// プロパティメソッド
Property Get time() As Date
  time = m_time
End Property

Property Get ip() As String
  ip = m_ip
End Property

Property Get mask() As Integer
  mask = m_mask
End Property

Property Get nwAddr() As String
  nwAddr = m_nwAddr
End Property

Property Get ResposeTime() As Integer
  ResposeTime = m_res
End Property

Property Let setTime(timeString As String)
  Dim year, month, day, hour, minute, second As String
  '// ex 20220625212059 -> 2022/06/25 21:20:59
  year = Left(timeString, 4)
  month = Mid(timeString, 5, 2)
  day = Mid(timeString, 7, 2)
  hour = Mid(timeString, 9, 2)
  minute = Mid(timeString, 11, 2)
  second = Mid(timeString, 13, 2)
  
  Dim tmp As String
  tmp = year & "/" & month & "/" & day & " " & hour & ":" & minute & ":" & second
  
  m_time = CDate(tmp)

End Property

Property Let setIpMask(cidrStr As String)
  Dim tmp() As String
  Dim ipstr, bstr As String
  Dim bits As Integer
  Dim oct1, oct2, oct3, oct4 As Integer
  
  tmp = Split(cidrStr, "/")
  ipstr = tmp(0)
  bstr = tmp(1)
  
  bits = CInt(bstr)
  
  tmp = Split(ipstr, ".")
  oct1 = CInt(tmp(0))
  oct2 = CInt(tmp(1))
  oct3 = CInt(tmp(2))
  oct4 = CInt(tmp(3))
  
  Dim mask As Integer
  mask = 255
  Dim shift As Integer
  shift = 0
  If 0 <= bits And bits <= 7 Then
    shift = 8 - bits
    oct1 = oct1 And mask * (2 ^ shift)
    oct2 = oct2 And 0
    oct3 = oct3 And 0
    oct4 = oct4 And 0
    
  ElseIf 8 <= bits And bits <= 15 Then
    shift = 8 - (bits Mod 8)
    oct1 = oct1 And mask
    oct2 = oct2 And mask * (2 ^ shift)
    oct3 = oct3 And 0
    oct4 = oct4 And 0
    
  ElseIf 16 <= bits And bits <= 23 Then
    shift = 8 - (bits Mod 16)
    oct1 = oct1 And mask
    oct2 = oct2 And mask
    oct3 = oct3 And mask * (2 ^ shift)
    oct4 = oct4 And 0
    
  ElseIf 24 <= bits And bits <= 32 Then
    shift = 8 - (bits Mod 24)
    oct1 = oct1 And mask
    oct2 = oct2 And mask
    oct3 = oct3 And mask
    oct4 = oct4 And mask * (2 ^ shift)
  End If
  
  m_ip = ipstr
  m_mask = bits
  m_nwAddr = CStr(oct1) & "." & CStr(oct2) & "." & CStr(oct3) & "." & CStr(oct4)
  
End Property

Property Let setResponseTime(response As String)
  If IsNumeric(response) Then
    m_res = CInt(response)
  Else
    '// タイムアウト発生時
    m_res = -1
  End If
End Property



