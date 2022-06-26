Attribute VB_Name = "test_module"
Option Explicit

Sub test_datetime()
  Dim dt1 As Date
  Dim dt2 As Date
  
  dt1 = DateValue("2022/06/25 13:00:05")
  dt2 = DateValue("2022/06/26 15:00:05")
  
  Debug.Print " "
  Debug.Print "DateValue) dt1=" & Format(dt1, "yyyy-mm-dd hh:mm:ss")
  Debug.Print "DateValue) dt2=" & Format(dt2, "yyyy-mm-dd hh:mm:ss")
  Debug.Print "DateValue) dt2-dt1=" & DateDiff("s", dt1, dt2)
  
  dt1 = CDate("2022/06/25 13:00:05")
  dt2 = CDate("2022/06/26 15:00:05")
  
  Debug.Print " "
  Debug.Print "CDate) dt1=" & Format(dt1, "yyyy-mm-dd hh:mm:ss")
  Debug.Print "CDate) dt2=" & Format(dt2, "yyyy-mm-dd hh:mm:ss")
  Debug.Print "CDate) dt2-dt1=" & DateDiff("s", dt1, dt2)
  
  Dim t1 As Variant
  Dim t2 As Variant

  t1 = TimeValue("2022/06/25 13:00:05")
  t2 = TimeValue("2022/06/26 15:00:05")

  Debug.Print " "
  Debug.Print "TimeValue) t1=" & Format(t1, "yyyy-mm-dd hh:mm:ss")
  Debug.Print "TimeValue) t2=" & Format(t2, "yyyy-mm-dd hh:mm:ss")
  Debug.Print "TimeValue) t2-t1=" & DateDiff("s", t1, t2)
  
  
End Sub

'//====================================
'// LogRecord�N���X�̃e�X�g
'//====================================
Sub test_LogRecord1()
  Dim log As logRecord
  
  Set log = New logRecord
  
  log.setTime = "20201019133124"
  Debug.Print log.time
  
  Set log = Nothing
End Sub

Sub test_LogRecord2()
  Dim log As logRecord
  
  Set log = New logRecord
  
  log.setIpMask = "10.20.30.1/16"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
  '//
  log.setIpMask = "255.255.255.255/0"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
  log.setIpMask = "255.255.255.255/1"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
  log.setIpMask = "255.255.255.255/7"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
  '//
  log.setIpMask = "255.255.255.255/8"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
  log.setIpMask = "255.255.255.255/9"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
  log.setIpMask = "255.255.255.255/15"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
  '//
  log.setIpMask = "255.255.255.255/16"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
  log.setIpMask = "255.255.255.255/17"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
  log.setIpMask = "255.255.255.255/23"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
  '//
  log.setIpMask = "255.255.255.255/24"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
  log.setIpMask = "255.255.255.255/25"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
  log.setIpMask = "255.255.255.255/32"
  Debug.Print "ip " & log.ip & " mask " & log.mask & " network " & log.nwAddr
    
  Set log = Nothing
End Sub

'//====================================
'// NodeObserver�N���X�̃e�X�g
'//====================================
'// �����O�o�b�t�@�̃e�X�g
Sub test_NodeObserver1()
  Dim no As NodeObserver
  
  Set no = New NodeObserver
  
  no.setResponse = 2
  Debug.Print "no.respUBound: " & no.respUBound & " next cursor: " & no.respCursor
  no.setResponse = 3
  Debug.Print "no.respUBound: " & no.respUBound & " next cursor: " & no.respCursor
  no.setResponse = 4
  Debug.Print "no.respUBound: " & no.respUBound & " next cursor: " & no.respCursor
  
  
  Set no = Nothing
End Sub

'// �̏ᔻ��̃e�X�g
Sub test_NodeObserver2()
  Dim log As logRecord: Set log = New logRecord
  Dim no As NodeObserver: Set no = New NodeObserver
  
  log.setTime = "20220625121500"
  log.setIpMask = "10.20.30.1/16"
  log.setResponseTime = 2
  no.setLogRecord = log
  
  log.setTime = "20220625121510"
  log.setResponseTime = "-"
  no.setLogRecord = log
  
  log.setTime = "20220625121520"
  log.setResponseTime = 5
  no.setLogRecord = log
  
  Debug.Print "ip: " & no.ip & " ��ԁF" & no.status & " �̏����[s]�F" & no.failureDulation
    
  Set no = Nothing
End Sub

'//====================================
'// main���W���[���̃e�X�g
'//====================================
Sub test_mainSuite()
  Debug.Print "***** START TEST SUITE *****"
  test_main1
  test_main2
  test_main2_2
  test_main3
  test_main3_2
  test_main4
  Debug.Print "***** END TEST SUITE *****"
End Sub

Sub test_main1()
  Dim csvName As String
  Dim out As Dictionary
  
  csvName = ThisWorkbook.Path & "\log1.csv"
  Set out = New Dictionary
  
  Dim N As Integer: N = 1 '// �A���^�C���A�E�g臒lN�i��j
  Dim m As Integer: m = 3 '// ���ω����̒��ߌv�Z�񐔂��i��j
  Dim t As Integer: t = 200 '// ���ω������Ԃ�臒l���ims�j
  
  AnalyzeLog CStr(csvName), out, N, m, t
  
  Debug.Print "--------------------------------------------------"
  Debug.Print csvName & " N=" & N & " m=" & m
  Debug.Print "--------------------------------------------------"
  Dim tmpKey As Variant
  For Each tmpKey In out
    Debug.Print "ip:" & out.Item(tmpKey).ip _
              & ", ���:" & out.Item(tmpKey).status _
              & ",  ���m:" & out.Item(tmpKey).failureStart _
              & ",  ���A:" & out.Item(tmpKey).failureEnd _
              & ",  �̏�[s]:" & out.Item(tmpKey).failureDulation _
              & ",  ����[ms]:" & out.Item(tmpKey).avrResponse _
              & ",  TO��:" & out.Item(tmpKey).timeoutCount
  Next
  Debug.Print ""
  Set out = Nothing
End Sub

Sub test_main2()
  Dim csvName As String
  Dim out As Dictionary
  
  csvName = ThisWorkbook.Path & "\log2.csv"
  Set out = New Dictionary
  
  Dim N As Integer: N = 1 '// �A���^�C���A�E�g臒lN�i��j
  Dim m As Integer: m = 3 '// ���ω����̒��ߌv�Z�񐔂��i��j
  Dim t As Integer: t = 200 '// ���ω������Ԃ�臒l���ims�j
  
  AnalyzeLog CStr(csvName), out, N, m, t
  
  Debug.Print "--------------------------------------------------"
  Debug.Print csvName & " N=" & N & " m=" & m
  Debug.Print "--------------------------------------------------"
  Dim tmpKey As Variant
  For Each tmpKey In out
    Debug.Print "ip:" & out.Item(tmpKey).ip _
              & ", ���:" & out.Item(tmpKey).status _
              & ",  ���m:" & out.Item(tmpKey).failureStart _
              & ",  ���A:" & out.Item(tmpKey).failureEnd _
              & ",  �̏�[s]:" & out.Item(tmpKey).failureDulation _
              & ",  ����[ms]:" & out.Item(tmpKey).avrResponse _
              & ",  TO��:" & out.Item(tmpKey).timeoutCount
  Next
  Debug.Print ""
  Set out = Nothing
End Sub

Sub test_main2_2()
  Dim csvName As String
  Dim out As Dictionary
  
  csvName = ThisWorkbook.Path & "\log2.csv"
  Set out = New Dictionary
  
  Dim N As Integer: N = 2 '// �A���^�C���A�E�g臒lN�i��j
  Dim m As Integer: m = 3 '// ���ω����̒��ߌv�Z�񐔂��i��j
  Dim t As Integer: t = 200 '// ���ω������Ԃ�臒l���ims�j
  
  AnalyzeLog CStr(csvName), out, N, m, t
  
  Debug.Print "--------------------------------------------------"
  Debug.Print csvName & " N=" & N & " m=" & m
  Debug.Print "--------------------------------------------------"
  Dim tmpKey As Variant
  For Each tmpKey In out
    Debug.Print "ip:" & out.Item(tmpKey).ip _
              & ", ���:" & out.Item(tmpKey).status _
              & ",  ���m:" & out.Item(tmpKey).failureStart _
              & ",  ���A:" & out.Item(tmpKey).failureEnd _
              & ",  �̏�[s]:" & out.Item(tmpKey).failureDulation _
              & ",  ����[ms]:" & out.Item(tmpKey).avrResponse _
              & ",  TO��:" & out.Item(tmpKey).timeoutCount
  Next
  Debug.Print ""
  Set out = Nothing
End Sub

Sub test_main3()
  Dim csvName As String
  Dim out As Dictionary
  
  csvName = ThisWorkbook.Path & "\log3.csv"
  Set out = New Dictionary
  
  Dim N As Integer: N = 1 '// �A���^�C���A�E�g臒lN�i��j
  Dim m As Integer: m = 3 '// ���ω����̒��ߌv�Z�񐔂��i��j
  Dim t As Integer: t = 200 '// ���ω������Ԃ�臒l���ims�j
  
  AnalyzeLog CStr(csvName), out, N, m, t
  
  Debug.Print "--------------------------------------------------"
  Debug.Print csvName & " N=" & N & " m=" & m
  Debug.Print "--------------------------------------------------"
  Dim tmpKey As Variant
  For Each tmpKey In out
    Debug.Print "ip:" & out.Item(tmpKey).ip _
              & ", ���:" & out.Item(tmpKey).status _
              & ",  ���m:" & out.Item(tmpKey).failureStart _
              & ",  ���A:" & out.Item(tmpKey).failureEnd _
              & ",  �̏�[s]:" & out.Item(tmpKey).failureDulation _
              & ",  ����[ms]:" & out.Item(tmpKey).avrResponse _
              & ",  TO��:" & out.Item(tmpKey).timeoutCount
  Next
  Debug.Print ""
  Set out = Nothing
End Sub

Sub test_main3_2()
  Dim csvName As String
  Dim out As Dictionary
  
  csvName = ThisWorkbook.Path & "\log3.csv"
  Set out = New Dictionary
  
  Dim N As Integer: N = 2 '// �A���^�C���A�E�g臒lN�i��j
  Dim m As Integer: m = 3 '// ���ω����̒��ߌv�Z�񐔂��i��j
  Dim t As Integer: t = 200 '// ���ω������Ԃ�臒l���ims�j
  
  AnalyzeLog CStr(csvName), out, N, m, t
  
  Debug.Print "--------------------------------------------------"
  Debug.Print csvName & " N=" & N & " m=" & m
  Debug.Print "--------------------------------------------------"
  Dim tmpKey As Variant
  For Each tmpKey In out
    Debug.Print "ip:" & out.Item(tmpKey).ip _
              & ", ���:" & out.Item(tmpKey).status _
              & ",  ���m:" & out.Item(tmpKey).failureStart _
              & ",  ���A:" & out.Item(tmpKey).failureEnd _
              & ",  �̏�[s]:" & out.Item(tmpKey).failureDulation _
              & ",  ����[ms]:" & out.Item(tmpKey).avrResponse _
              & ",  TO��:" & out.Item(tmpKey).timeoutCount
  Next
  Debug.Print ""
  Set out = Nothing
End Sub

Sub test_main4()
  Dim csvName As String
  Dim out As Dictionary
  
  csvName = ThisWorkbook.Path & "\log4.csv"
  Set out = New Dictionary
  
  Dim N As Integer: N = 2 '// �A���^�C���A�E�g臒lN�i��j
  Dim m As Integer: m = 3 '// ���ω����̒��ߌv�Z�񐔂��i��j
  Dim t As Integer: t = 200 '// ���ω������Ԃ�臒l���ims�j
  
  AnalyzeLog CStr(csvName), out, N, m, t
  
  Debug.Print "--------------------------------------------------"
  Debug.Print csvName & " N=" & N & " m=" & m
  Debug.Print "--------------------------------------------------"
  Dim tmpKey As Variant
  For Each tmpKey In out
    Debug.Print "ip:" & out.Item(tmpKey).ip _
              & ", ���:" & out.Item(tmpKey).status _
              & ",  ���m:" & out.Item(tmpKey).failureStart _
              & ",  ���A:" & out.Item(tmpKey).failureEnd _
              & ",  �̏�[s]:" & out.Item(tmpKey).failureDulation _
              & ",  ����[ms]:" & out.Item(tmpKey).avrResponse _
              & ",  TO��:" & out.Item(tmpKey).timeoutCount
  Next
  Debug.Print ""
  Set out = Nothing
End Sub
