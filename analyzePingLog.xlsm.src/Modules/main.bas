Attribute VB_Name = "main"
Option Explicit

Sub UI_ログファイル開く()
  Dim csvName As Variant
  csvName = Application.GetOpenFilename(FileFilter:="CSVファイル(*.csv),*.csv", _
                                        Title:="CSVファイルの選択")
  Dim out As Dictionary
  Set out = New Dictionary
  
  Dim wsSetting As Worksheet
  Set wsSetting = Worksheets("SETTING")
  Dim N As Integer '// 連続タイムアウト閾値N（回）
  Dim m As Integer '// 平均応答の直近計算回数ｍ（回）
  Dim t As Integer '// 平均応答時間の閾値ｔ（ms）
  
  N = wsSetting.Range("B6").Value
  m = wsSetting.Range("B7").Value
  t = wsSetting.Range("B8").Value
  
  AnalyzeLog CStr(csvName), out, N, m, t
  
  Dim ows As Worksheet
  Worksheets.Add After:=Sheets(Worksheets.Count)
  Set ows = Worksheets(Worksheets.Count)
    
  ows.Range("A1").Value = "IP"
  ows.Range("B1").Value = "状態"
  ows.Range("C1").Value = "故障期間[s]"
  ows.Range("D1").Value = "検知時刻"
  ows.Range("E1").Value = "復帰時刻"
  ows.Range("F1").Value = "平均応答時間[s]"
  
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
'　メイン処理
'-----------------------
' [IN ] csvName CSVファイル名
' [OUT] outDict 監視ログ解析結果（辞書型）
' [IN ] th      連続タイムアウト閾値N（回）
' [IN ] avrC    平均応答の直近計算回数ｍ（回）
' [IN ] avrT    平均応答時間の閾値ｔ（ms）（未使用）
'
Sub AnalyzeLog(csvName As String, outDict As Dictionary, th As Integer, avrC As Integer, avrT As Integer)
  
  '-----------------------
  '　CSV読込み
  '-----------------------
  If csvName = "" Then
    MsgBox "ファイル名を入力してください"
    Exit Sub
  End If
  
  '// 最終行にシークして行数を取り出す
  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  Dim tso As TextStream
  Dim lineCount As Long
  
  Set tso = fso.OpenTextFile(csvName, ForAppending)

  lineCount = tso.Line - 1
  tso.Close
  
  '-----------------------
  '　ログのデータ化
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
        '// 時刻
        log(i).setTime = strLogCols(j)
      ElseIf j = 1 Then
        '// CIDR
        log(i).setIpMask = strLogCols(j)
        If Not ipList.Exists(log(i).ip) Then
          '// 新規IP登録、データ数=1で初期化
          ipList.Add log(i).ip, 1
        Else
          '// データ数更新
          ipList.Item(log(i).ip) = ipList.Item(log(i).ip) + 1
        End If
      ElseIf j = 2 Then
        '// Response
        log(i).setResponseTime = strLogCols(j)
      End If
    Next
  Next
  
'  '// IPリスト表示
'  Dim tmp As Variant
'  For Each tmp In ipList
'    Debug.Print "key:" & tmp & " Value:" & ipList.Item(tmp)
'  Next
  
  '-----------------------
  '　動作パラメータの取得
  '-----------------------
'  Dim wsSetting As Worksheet
'  Set wsSetting = Worksheets("SETTING")
'  Dim th As Integer   '// 連続タイムアウト閾値N（回）
'  Dim avrC As Integer '// 平均応答の直近計算回数ｍ（回）
'  Dim avrT As Integer '// 平均応答時間の閾値ｔ（ms）
'
'  th = wsSetting.Range("B6").Value
'  avrC = wsSetting.Range("B7").Value
'  avrT = wsSetting.Range("B8").Value
  
  '-----------------------
  '　IP毎の故障チェック
  '-----------------------
  Dim tmpKey As Variant
  Dim rec As Variant
  Dim arrLog As Variant
  Dim noDict As Dictionary
  Dim noObj As NodeObserver
  
  Set noDict = New Dictionary
 
  '// 監視テーブル初期化
  For Each tmpKey In ipList
    Set noObj = New NodeObserver
    noObj.setTimeoutThreshold = th
    noObj.setAvrCount = avrC
    noObj.setAvrTimeout = avrT
    noDict.Add tmpKey, noObj
  Next
  
  '// 監視テーブルを最新ログで更新
  For Each tmpKey In ipList
    For i = 1 To UBound(log)
      If tmpKey = log(i).ip Then
        noDict.Item(tmpKey).setLogRecord = log(i)
      End If
    Next
  Next
  
'  '// 監視テーブル表示（Debug）
'  For Each tmpKey In noDict
'    Debug.Print "ip:" & noDict.Item(tmpKey).ip _
'              & " 状態:" & noDict.Item(tmpKey).status _
'              & " 故障期間[s]" & noDict.Item(tmpKey).failureDulation _
'              & " 平均応答[ms]" & noDict.Item(tmpKey).avrResponse
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












