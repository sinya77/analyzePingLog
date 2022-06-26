Attribute VB_Name = "main"
Option Explicit

Sub UI_���O�t�@�C���J��()
  Dim csvName As Variant
  csvName = Application.GetOpenFilename(FileFilter:="CSV�t�@�C��(*.csv),*.csv", _
                                        Title:="CSV�t�@�C���̑I��")
  Dim out As Dictionary
  Set out = New Dictionary
  
  Dim wsSetting As Worksheet
  Set wsSetting = Worksheets("SETTING")
  Dim N As Integer '// �A���^�C���A�E�g臒lN�i��j
  Dim m As Integer '// ���ω����̒��ߌv�Z�񐔂��i��j
  Dim t As Integer '// ���ω������Ԃ�臒l���ims�j
  
  N = wsSetting.Range("B6").Value
  m = wsSetting.Range("B7").Value
  t = wsSetting.Range("B8").Value
  
  AnalyzeLog CStr(csvName), out, N, m, t
  
  Dim ows As Worksheet
  Worksheets.Add After:=Sheets(Worksheets.Count)
  Set ows = Worksheets(Worksheets.Count)
    
  ows.Range("A1").Value = "IP"
  ows.Range("B1").Value = "���"
  ows.Range("C1").Value = "�̏����[s]"
  ows.Range("D1").Value = "���m����"
  ows.Range("E1").Value = "���A����"
  ows.Range("F1").Value = "���ω�������[s]"
  
  '//
  With ows
    .Range("D:E").NumberFormatLocal = "yyyy/MM/dd hh:mm:ss"
    .Select
  End With
  
  Dim tmpKey As Variant
  Dim i As Integer
  i = 2
  For Each tmpKey In out
    ows.Range("A" & i).Value = out.Item(tmpKey).ip
    ows.Range("B" & i).Value = out.Item(tmpKey).status
    ows.Range("C" & i).Value = out.Item(tmpKey).failureDulation
    ows.Range("D" & i).Value = out.Item(tmpKey).failureStart
    ows.Range("E" & i).Value = out.Item(tmpKey).failureEnd
    ows.Range("F" & i).Value = out.Item(tmpKey).avrResponse
    i = i + 1
  Next
  
  Set out = Nothing
  
End Sub

'-----------------------
'�@���C������
'-----------------------
' [IN ] csvName CSV�t�@�C����
' [OUT] outDict �Ď����O��͌��ʁi�����^�j
' [IN ] th      �A���^�C���A�E�g臒lN�i��j
' [IN ] avrC    ���ω����̒��ߌv�Z�񐔂��i��j
' [IN ] avrT    ���ω������Ԃ�臒l���ims�j�i���g�p�j
'
Sub AnalyzeLog(csvName As String, outDict As Dictionary, th As Integer, avrC As Integer, avrT As Integer)
  
  '-----------------------
  '�@CSV�Ǎ���
  '-----------------------
  If csvName = "" Then
    MsgBox "�t�@�C��������͂��Ă�������"
    Exit Sub
  End If
  
  '// �ŏI�s�ɃV�[�N���čs�������o��
  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  Dim tso As TextStream
  Dim lineCount As Long
  
  Set tso = fso.OpenTextFile(csvName, ForAppending)

  lineCount = tso.Line - 1
  tso.Close
  
  '-----------------------
  '�@���O�̃f�[�^��
  '-----------------------
  Set tso = fso.OpenTextFile(csvName, ForReading)
  
  Dim i, j As Long
  Dim strLogRow As String
  Dim strLogCols() As String
    
  Dim log() As logRecord
  ReDim log(lineCount)
  Dim ipList As Dictionary:     Set ipList = New Dictionary
  '//Dim nwAddrList As Dictionary: Set nwAddrList = New Dictionary
  
  For i = 1 To lineCount
    strLogRow = tso.ReadLine
    '//Debug.Print strLogRow
    strLogCols = Split(strLogRow, ",")
    Set log(i) = New logRecord
    
    For j = 0 To UBound(strLogCols)
      '//Debug.Print "(" & i & "," & j & ")" & strLogCols(j)
      If j = 0 Then
        '// ����
        log(i).setTime = strLogCols(j)
      ElseIf j = 1 Then
        '// CIDR
        log(i).setIpMask = strLogCols(j)
        If Not ipList.Exists(log(i).ip) Then
          '// �V�KIP�o�^�A�f�[�^��=1�ŏ�����
          ipList.Add log(i).ip, 1
        Else
          '// �f�[�^���X�V
          ipList.Item(log(i).ip) = ipList.Item(log(i).ip) + 1
        End If
      ElseIf j = 2 Then
        '// Response
        log(i).setResponseTime = strLogCols(j)
      End If
    Next
  Next
  
'  '// IP���X�g�\��
'  Dim tmp As Variant
'  For Each tmp In ipList
'    Debug.Print "key:" & tmp & " Value:" & ipList.Item(tmp)
'  Next
  
  '-----------------------
  '�@����p�����[�^�̎擾
  '-----------------------
'  Dim wsSetting As Worksheet
'  Set wsSetting = Worksheets("SETTING")
'  Dim th As Integer   '// �A���^�C���A�E�g臒lN�i��j
'  Dim avrC As Integer '// ���ω����̒��ߌv�Z�񐔂��i��j
'  Dim avrT As Integer '// ���ω������Ԃ�臒l���ims�j
'
'  th = wsSetting.Range("B6").Value
'  avrC = wsSetting.Range("B7").Value
'  avrT = wsSetting.Range("B8").Value
  
  '-----------------------
  '�@IP���̌̏�`�F�b�N
  '-----------------------
  Dim tmpKey As Variant
  Dim rec As Variant
  Dim arrLog As Variant
  Dim noDict As Dictionary
  Dim noObj As NodeObserver
  
  Set noDict = New Dictionary
 
  '// �Ď��e�[�u��������
  For Each tmpKey In ipList
    Set noObj = New NodeObserver
    noObj.setTimeoutThreshold = th
    noObj.setAvrCount = avrC
    noObj.setAvrTimeout = avrT
    noDict.Add tmpKey, noObj
  Next
  
  '// �Ď��e�[�u�����ŐV���O�ōX�V
  For Each tmpKey In ipList
    For i = 1 To UBound(log)
      If tmpKey = log(i).ip Then
        noDict.Item(tmpKey).setLogRecord = log(i)
      End If
    Next
  Next
  
'  '// �Ď��e�[�u���\���iDebug�j
'  For Each tmpKey In noDict
'    Debug.Print "ip:" & noDict.Item(tmpKey).ip _
'              & " ���:" & noDict.Item(tmpKey).status _
'              & " �̏����[s]" & noDict.Item(tmpKey).failureDulation _
'              & " ���ω���[ms]" & noDict.Item(tmpKey).avrResponse
'  Next
  
  Set outDict = noDict
  
  '//
  Set ipList = Nothing
  '//Set nwAddrList = Nothing
    
  For i = 1 To lineCount
    Set log(i) = Nothing
  Next
  Set ipList = Nothing
  Set noDict = Nothing
    
  Set fso = Nothing
  Set tso = Nothing
    
  
End Sub












