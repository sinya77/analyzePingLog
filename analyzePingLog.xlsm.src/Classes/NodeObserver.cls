VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NodeObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// �c�[�� -> �I�v�V���� -> �S�� �̃G���[�g���b�v��
'// �u�N���X ���W���[���Œ��f�v

Option Explicit

'// �v���p�e�B
Private m_ip As String
Private m_respTimeRingBuffer() As Integer
Private m_failureStart As Date
Private m_failureEnd As Date
Private m_failureDulation As Long
Private m_status As Integer         '// -1=�f�[�^�Ȃ�, 0=����, 1=�̏�
Private m_timeoutCount As Integer   '// �A���^�C���A�E�g��
Private m_timeoutThreshold As Integer '// �A���^�C���A�E�g臒l

'//�@�ݒ�v���p�e�B
Private m_avrTimeout As Integer     '// ���ω������Ԃ�臒l
Private m_avrCount As Integer       '// ���ω����̒��ߌv�Z��

'//�@�W�v�l�v���p�e�B
Private m_averageResponse As Integer
Private m_ringBufIndex As Integer     '//
Private m_ringBufPrev As Integer


' ����������
Private Sub Class_Initialize()
  m_ip = "0.0.0.0"
  m_failureStart = CDate("2000/01/01 00:00:00")
  m_failureEnd = CDate("2000/01/01 00:00:00")
  m_failureDulation = 0
  '//
  m_avrTimeout = 100          '// �f�t�H���g�l
  m_avrCount = 4              '// �f�t�H���g�l
  m_timeoutThreshold = 1      '// �f�t�H���g�l
  
  '// ���������o�b�t�@�̏�����
  ReDim m_respTimeRingBuffer(1 To m_avrCount) '// ���߂���̃f�[�^��ۑ�
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

' �I������
Private Sub Class_Terminate()
  'Debug.Print "terminate ip:" & m_ip
End Sub

'// �v���p�e�B���\�b�h
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

'//�@����p�����[�^���Z�b�g
Property Let setTimeoutThreshold(ByVal timeoutThreshold As Integer)
  m_timeoutThreshold = timeoutThreshold
End Property

Property Let setAvrTimeout(ByVal avrTimeout As Integer)
  m_avrTimeout = avrTimeout
End Property

Property Let setAvrCount(ByVal avrCount As Integer)
  m_avrCount = avrCount
  
  '// ���������o�b�t�@�̏�����
  ReDim m_respTimeRingBuffer(1 To m_avrCount) '// ���߂���̃f�[�^��ۑ�
  Dim i As Integer
  For i = 1 To UBound(m_respTimeRingBuffer)
    m_respTimeRingBuffer(i) = -1
  Next
  
End Property

'//�@���O���R�[�h�I�u�W�F�N�g���󂯎��Ď������X�V
Property Let setLogRecord(log As logRecord)
  m_ip = log.ip
  Me.setResponse = log.ResposeTime
  Call m_calcAvrResponse
  
  '//----------------------------------------------------------------------
  If m_status = -1 Then                          '// ��ԁ��f�[�^�Ȃ�
    If m_respTimeRingBuffer(m_ringBufPrev) <> -1 Then
      m_status = 0  '//�����
    Else
      m_failureStart = log.time                     '// �̏ጟ�m�����Z�b�g
      m_timeoutCount = m_timeoutCount + 1           '// �^�C���A�E�g�J�E���g�X�V
      If m_timeoutCount >= m_timeoutThreshold Then
        m_status = 1  '//�̏��
      Else
        m_status = 0  '//�����
      End If
    End If
  '//----------------------------------------------------------------------
  ElseIf m_status = 0 Then                       '// ��ԁ�����
    If m_respTimeRingBuffer(m_ringBufPrev) <> -1 Then
      m_status = 0  '//�����
    Else
      m_failureStart = log.time                     '// �̏ጟ�m�����Z�b�g
      m_timeoutCount = m_timeoutCount + 1           '// �^�C���A�E�g�J�E���g�X�V
      If m_timeoutCount >= m_timeoutThreshold Then
        m_status = 1  '//�̏��
      Else
        m_status = 0  '//�����
      End If
    End If
  '//----------------------------------------------------------------------
  ElseIf m_status = 1 Then                       '// ��ԁ��̏�
    If m_respTimeRingBuffer(m_ringBufPrev) <> -1 Then
      m_status = 0  '//�����
      m_timeoutCount = 0
      m_failureEnd = log.time                     '// �̏ᕜ�A�����Z�b�g
      m_failureDulation = DateDiff("s", m_failureStart, m_failureEnd)
    Else
      m_status = 1  '//�̏��
      '//m_timeoutCount = 0
    End If
  End If
  '//----------------------------------------------------------------------
  
End Property

'// ���߂���̉����������O�o�b�t�@�Ɋi�[
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

'// �����O�o�b�t�@�̏���擾
Property Get respUBound() As Integer
  respUBound = UBound(m_respTimeRingBuffer)
End Property

'// �����O�o�b�t�@�̃J�[�\���擾
Property Get respCursor() As Integer
  respCursor = m_ringBufIndex
End Property
'// ���ω������Ԃ̌v�Z
'// ���l�F�^�C���A�E�g�i"-"�j�͌v�Z���珜�O
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


