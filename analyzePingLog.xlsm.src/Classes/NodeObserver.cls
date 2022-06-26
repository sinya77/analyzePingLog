VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NodeObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// ツール -> オプション -> 全般 のエラートラップを
'// 「クラス モジュールで中断」

Option Explicit

'// プロパティ
Private m_ip As String
Private m_respTimeRingBuffer() As Integer
Private m_failureStart As Date
Private m_failureEnd As Date
Private m_failureDulation As Long
Private m_status As Integer         '// -1=データなし, 0=正常, 1=故障
Private m_timeoutCount As Integer   '// 連続タイムアウト回数
Private m_timeoutThreshold As Integer '// 連続タイムアウト閾値

'//　設定プロパティ
Private m_avrTimeout As Integer     '// 平均応答時間の閾値
Private m_avrCount As Integer       '// 平均応答の直近計算回数

'//　集計値プロパティ
Private m_averageResponse As Integer
Private m_ringBufIndex As Integer     '//
Private m_ringBufPrev As Integer


' 初期化処理
Private Sub Class_Initialize()
  m_ip = "0.0.0.0"
  m_failureStart = CDate("2000/01/01 00:00:00")
  m_failureEnd = CDate("2000/01/01 00:00:00")
  m_failureDulation = 0
  '//
  m_avrTimeout = 100          '// デフォルト値
  m_avrCount = 4              '// デフォルト値
  m_timeoutThreshold = 1      '// デフォルト値
  
  '// 応答履歴バッファの初期化
  ReDim m_respTimeRingBuffer(1 To m_avrCount) '// 直近ｍ回のデータを保存
  Dim i As Integer
  For i = 1 To UBound(m_respTimeRingBuffer)
    m_respTimeRingBuffer(i) = -1
  Next
  '// Debug
  Call m_calcAvrResponse
  '// Debug
  
  m_ringBufIndex = 1
  m_ringBufPrev = 1
  m_status = -1
  m_timeoutCount = 0
End Sub

' 終了処理
Private Sub Class_Terminate()
  'Debug.Print "terminate ip:" & m_ip
End Sub

'// プロパティメソッド
Property Get ip() As String
  ip = m_ip
End Property

Property Get failureStart() As Date
  failureStart = m_failureStart
End Property

Property Get failureEnd() As Date
  failureEnd = m_failureEnd
End Property

Property Get failureDulation() As Long
  failureDulation = m_failureDulation
End Property

Property Get status() As Integer
  status = m_status
End Property

Property Get avrResponse() As Integer
  avrResponse = m_averageResponse
End Property

Property Get timeoutCount() As Integer
  timeoutCount = m_timeoutCount
End Property

'//　動作パラメータをセット
Property Let setTimeoutThreshold(ByVal timeoutThreshold As Integer)
  m_timeoutThreshold = timeoutThreshold
End Property

Property Let setAvrTimeout(ByVal avrTimeout As Integer)
  m_avrTimeout = avrTimeout
End Property

Property Let setAvrCount(ByVal avrCount As Integer)
  m_avrCount = avrCount
  
  '// 応答履歴バッファの初期化
  ReDim m_respTimeRingBuffer(1 To m_avrCount) '// 直近ｍ回のデータを保存
  Dim i As Integer
  For i = 1 To UBound(m_respTimeRingBuffer)
    m_respTimeRingBuffer(i) = -1
  Next
  
End Property

'//　ログレコードオブジェクトを受け取り監視情報を更新
Property Let setLogRecord(log As logRecord)
  m_ip = log.ip
  Me.setResponse = log.ResposeTime
  Call m_calcAvrResponse
  
  '//----------------------------------------------------------------------
  If m_status = -1 Then                          '// 状態＝データなし
    If m_respTimeRingBuffer(m_ringBufPrev) <> -1 Then
      m_status = 0  '//正常へ
    Else
      m_failureStart = log.time                     '// 故障検知時刻セット
      m_timeoutCount = m_timeoutCount + 1           '// タイムアウトカウント更新
      If m_timeoutCount >= m_timeoutThreshold Then
        m_status = 1  '//故障へ
      Else
        m_status = 0  '//正常へ
      End If
    End If
  '//----------------------------------------------------------------------
  ElseIf m_status = 0 Then                       '// 状態＝正常
    If m_respTimeRingBuffer(m_ringBufPrev) <> -1 Then
      m_status = 0  '//正常へ
    Else
      m_failureStart = log.time                     '// 故障検知時刻セット
      m_timeoutCount = m_timeoutCount + 1           '// タイムアウトカウント更新
      If m_timeoutCount >= m_timeoutThreshold Then
        m_status = 1  '//故障へ
      Else
        m_status = 0  '//正常へ
      End If
    End If
  '//----------------------------------------------------------------------
  ElseIf m_status = 1 Then                       '// 状態＝故障
    If m_respTimeRingBuffer(m_ringBufPrev) <> -1 Then
      m_status = 0  '//正常へ
      m_timeoutCount = 0
      m_failureEnd = log.time                     '// 故障復帰時刻セット
      m_failureDulation = DateDiff("s", m_failureStart, m_failureEnd)
    Else
      m_status = 1  '//故障へ
      '//m_timeoutCount = 0
    End If
  End If
  '//----------------------------------------------------------------------
  
End Property

'// 直近ｍ回の応答をリングバッファに格納
Property Let setResponse(respTime As Integer)
  m_ringBufPrev = m_ringBufIndex
  '//
  m_respTimeRingBuffer(m_ringBufIndex) = respTime
  If m_ringBufIndex < UBound(m_respTimeRingBuffer) Then
    m_ringBufIndex = m_ringBufIndex + 1
  Else
    m_ringBufIndex = 1
  End If
End Property

'// リングバッファの上限取得
Property Get respUBound() As Integer
  respUBound = UBound(m_respTimeRingBuffer)
End Property

'// リングバッファのカーソル取得
Property Get respCursor() As Integer
  respCursor = m_ringBufIndex
End Property
'// 平均応答時間の計算
'// 備考：タイムアウト（"-"）は計算から除外
Private Function m_calcAvrResponse()
  Dim sum, cnt As Integer
  sum = 0
  cnt = 0
  Dim i As Integer
  For i = 1 To UBound(m_respTimeRingBuffer)
    'Debug.Print "ip:" & m_ip & " resp(" & i & ") " & m_respTimeRingBuffer(i)
    If m_respTimeRingBuffer(i) <> -1 Then
      sum = sum + m_respTimeRingBuffer(i)
      cnt = cnt + 1
    End If
  Next
  If cnt <> 0 Then
    m_averageResponse = sum / cnt
  Else
    m_averageResponse = 0
  End If
  'Debug.Print "ip:" & m_ip & " resp sum:" & sum & " cnt:" & cnt & " avr:" & m_averageResponse
End Function



