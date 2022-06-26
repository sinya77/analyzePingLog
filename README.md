# analyzePingLog
ping応答ログを解析するVBA

## ログフォーマット
`＜確認日時＞,＜サーバアドレス＞,＜応答結果＞`
上記を1行とするCSVファイルである。
確認日時は、YYYYMMDDhhmmssの形式。ただし、年＝YYYY（4桁の数字）、月＝MM（2桁の数字。以下同様）、日＝DD、時＝hh、分＝mm、秒＝ssである。
サーバアドレスは、ネットワークプレフィックス長付きのIPv4アドレスである。
応答結果には、pingの応答時間がミリ秒単位で記載される。ただし、タイムアウトした場合は"-"(ハイフン記号)となる。

## 機能
### 故障判定
pingがタイムアウトした場合を故障とみなし、サーバの故障時間を表示する。  
最初にタイムアウトしたときから、次にpingの応答が返るまでを故障期間とする。  
サーバアドレスごとに判定する。

### タイムアウト回数設定
N回以上連続してタイムアウトした場合にのみ故障とみなす。  
Nは設定。

### 平均応答時間
直近ｍ回のPing応答平均応答時間を算出する。  
直近ｍ回にタイムアウト（"-"、ハイフン）が含まれる場合は除外して計算する。  
ｍは設定。

### 過負荷判定（将来予定）
直近ｍ回のPing応答平均応答時間がtミリ秒を超えた場合に過負荷とみなす。

### スイッチ障害検知（将来予定）
サブネット内のサーバがすべて故障と判定された場合、サブネット（スイッチ）の呼称と見なし、サブネットの故障時間を表示する。

# 使用方法
1. analyzePingLog.xlsmの「SETTING」シートを設定する
2. 「ログファイルを開く」をクリックし、ログファイルを選択する。
3. 新規シートが追加され、結果が出力される。

（補足）  
Excelマクロの以下の参照設定を有効にしてください。  
Excelメインメニュー ＞ 開発 ＞ Visual Basic、を選択  
VBEエディタのツール ＞ 参照設定にて以下を設定  
・Microsoft Scripting Runtime

## 出力シートの見方
- IP： ログに記録されたサーバアドレス
- 状態：サーバの状態。0=正常、1=故障。
- 検知時刻：タイムアウトを1回以上検知した時刻。
- 復帰時刻：「故障」状態から1回以上Ping応答を受信した時刻。このとき状態が0（正常）に変わる。
- 平均応答時間：Ping応答平均応答時間

### 出力例1
**IP**|**状態**|**故障期間[s]**|**検知時刻**|**復帰時刻**|**平均応答時間[s]**
:-----:|:-----:|:-----:|:-----:|:-----:|:-----:
10.20.30.1|0|60|2020/10/19 13:32:24|2020/10/19 13:33:24|4

サーバ（10.20.30.1）はタイムアウト後に復旧しており、故障時間は60秒。直近ｍ回の平均応答時間は4ms

### 出力例2
**IP**|**状態**|**故障期間[s]**|**検知時刻**|**復帰時刻**|**平均応答時間[s]**
:-----:|:-----:|:-----:|:-----:|:-----:|:-----:
10.20.30.1|1|0|2020/10/19 13:32:24|2000/01/01 00:00:00|2

N=1の設定。サーバ（10.20.30.1）はタイムアウトにより故障状態（状態=1）。復旧していないため復帰時刻は初期値が表示される。

### 出力例3
**IP**|**状態**|**故障期間[s]**|**検知時刻**|**復帰時刻**|**平均応答時間[s]**
:-----:|:-----:|:-----:|:-----:|:-----:|:-----:
10.20.30.1|0|0|2020/10/19 13:32:24|2000/01/01 00:00:00|2

N=2の設定。サーバ（10.20.30.1）はタイムアウトを1回検知のため正常状態（状態=0）。

# プログラム構造
- analyzePingLog.xlsm　　　：　マクロブック本体
- analyzePingLog.xlsm.src　：　マクロエクスポートファイル
  - Classes
    - LogRecord.cls　　　　：ログレコード管理クラス
    - NodeObserver.cls　　：サーバ監視管理クラス
  - Modules
    - main.bas             ：メインモジュール
    - test_module.bas      ：ユニットテスト用モジュール
- log1.csv ：テストデータ
- log2.csv ：テストデータ
- log3.csv ：テストデータ
- log4.csv ：テストデータ

## 処理の流れ
1. マクロボタンからUI_ログファイル開く()を呼び出し
1. CSVファイルのパス取得、SETTINGシートから設定を読み出し、ログ解析本体（AnalyzeLog）を呼び出し
1. AnalyzeLogはCSVログファイルをログレコード管理クラスを使ってデータ化し、時刻、サーバアドレス、サブネットマスク、応答時間を内部形式でメモリ上に保存。
1. 次にサーバ監視管理クラスを使ってサーバアドレス毎にログデータをCSVファイルの記録順に振り分ける。
    1. このときサーバ管理クラスでは直近ｍ回分の応答時間を内部に溜め、平均応答時間を算出する。 
    1. またタイムアウト回数設定に従って状態判定、タイムアウト検知時刻、故障復帰時刻を記録する
1．AnalyzeLogは呼出し元へサーバ管理クラスのコピーを返す
1．呼出し元（UI_ログファイル開く）にて解析結果を新規シートに出力する。

# 検査
## 検査内容
log.1.csv（スモークテスト、提示されたサンプルデータ）、設定N=1
以下の期待値は見やすいようにサーバアドレス毎にグルーピング

```
20201019133124,10.20.30.1/16,2
20201019133224,10.20.30.1/16,522
20201019133324,10.20.30.1/16,-
期待値）状態=0、開始=2020/10/19 13:33:24、復帰=2000/01/01、故障期間=0、平均=262ms、TO回数=1

20201019133125,10.20.30.2/16,1
20201019133225,10.20.30.2/16,1
20201019133325,10.20.30.2/16,2
期待値）状態=0、開始=2000/01/01、復帰=2000/01/01、故障期間=0、平均=1ms、TO回数=0

20201019133134,192.168.1.1/24,10
20201019133234,192.168.1.1/24,8
期待値）状態=0、開始=2000/01/01、復帰=2000/01/01、故障期間=0、平均=9ms、TO回数=0

20201019133135,192.168.1.2/24,5
20201019133235,192.168.1.2/24,15
期待値）状態=0、開始=2000/01/01、復帰=2000/01/01、故障期間=0、平均=10ms、TO回数=0

```

log2.csv（IP1＝10.20.30.1）
```
・IP1がN=1、1回タイムアウト⇒復旧　（期待値）状態=0、検知、復旧、故障時間が記録されること
・IP1がN=2、1回タイムアウト⇒復旧	（期待値）状態=0、検知時刻のみ記録されること
```
log3.csv（IP1＝10.20.30.1）
```
・IP1がN=1、1回タイムアウトのみ		（期待値）状態=1、検知時刻のみが記録されること
・IP1がN=2、1回タイムアウトのみ		（期待値）状態=0、検知時刻のみ記録されること
```
log4.csv（IP1＝10.20.30.1、IP2＝192.168.1.1）
```
・IP1がN=2、2回タイムアウト⇒復旧	（期待値）状態=0、検知、復旧、故障時間が記録されること
・IP2がN=2、2回タイムアウトのみ		（期待値）状態=1、開始時刻のみ記録されること
```

## 検査結果
test_moduleのtest_mainSuite()実行結果を以下に示す。上記期待値通りになっていることを確認。
```
***** START TEST SUITE *****
--------------------------------------------------
log1.csv N=1 m=3
--------------------------------------------------
ip:10.20.30.1, 状態:1,  検知:2020/10/19 13:33:24,  復帰:2000/01/01,  故障[s]:0,  平均[ms]:262,  TO回数:1
ip:10.20.30.2, 状態:0,  検知:2000/01/01,  復帰:2000/01/01,  故障[s]:0,  平均[ms]:1,  TO回数:0
ip:192.168.1.1, 状態:0,  検知:2000/01/01,  復帰:2000/01/01,  故障[s]:0,  平均[ms]:9,  TO回数:0
ip:192.168.1.2, 状態:0,  検知:2000/01/01,  復帰:2000/01/01,  故障[s]:0,  平均[ms]:10,  TO回数:0

--------------------------------------------------
log2.csv N=1 m=3
--------------------------------------------------
ip:10.20.30.1, 状態:0,  検知:2020/10/19 13:32:24,  復帰:2020/10/19 13:33:24,  故障[s]:60,  平均[ms]:4,  TO回数:0
ip:192.168.1.1, 状態:0,  検知:2000/01/01,  復帰:2000/01/01,  故障[s]:0,  平均[ms]:9,  TO回数:0

--------------------------------------------------
log2.csv N=2 m=3
--------------------------------------------------
ip:10.20.30.1, 状態:0,  検知:2020/10/19 13:32:24,  復帰:2000/01/01,  故障[s]:0,  平均[ms]:4,  TO回数:1
ip:192.168.1.1, 状態:0,  検知:2000/01/01,  復帰:2000/01/01,  故障[s]:0,  平均[ms]:9,  TO回数:0

--------------------------------------------------
log3.csv N=1 m=3
--------------------------------------------------
ip:10.20.30.1, 状態:1,  検知:2020/10/19 13:32:24,  復帰:2000/01/01,  故障[s]:0,  平均[ms]:2,  TO回数:1
ip:192.168.1.1, 状態:0,  検知:2000/01/01,  復帰:2000/01/01,  故障[s]:0,  平均[ms]:9,  TO回数:0

--------------------------------------------------
log3.csv N=2 m=3
--------------------------------------------------
ip:10.20.30.1, 状態:0,  検知:2020/10/19 13:32:24,  復帰:2000/01/01,  故障[s]:0,  平均[ms]:2,  TO回数:1
ip:192.168.1.1, 状態:0,  検知:2000/01/01,  復帰:2000/01/01,  故障[s]:0,  平均[ms]:9,  TO回数:0

--------------------------------------------------
log4.csv N=2 m=3
--------------------------------------------------
ip:10.20.30.1, 状態:0,  検知:2020/10/19 13:33:24,  復帰:2020/10/19 13:34:24,  故障[s]:60,  平均[ms]:10,  TO回数:0
ip:192.168.1.1, 状態:1,  検知:2020/10/19 13:34:54,  復帰:2000/01/01,  故障[s]:0,  平均[ms]:8,  TO回数:2

***** END TEST SUITE *****
```
