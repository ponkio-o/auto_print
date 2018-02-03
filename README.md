# Touch on Time Auto_print
勤怠管理ソフトTouch On Timeの日別データを自動的に印刷します。  
また日別の勤務時間などを自動的に取得してxslx形式で保存します。

# 使い方
`auto_print.exe`をクリックするだけで使えます。

## 必要なもの
- GoogleChrome
- [chromedriver](https://sites.google.com/a/chromium.org/chromedriver/downloads)
- MicrosoftOffice

Officeは必須ではないですが、勤怠データをxlsx形式で保存するため閲覧するのに必要です。   
事前にchromedriverを`auto_print.exe`と同一ディレクトリに設置してください。

## 設定について
設定は基本的に`config.ini`より行います。

```
[shop_settings]
username  = user         #ログインID
password  = password     #ログインPW
shop_id   = 100000000    #所属(HTMLよりVALUEを調べる)

[crawl_settings]
start_day = 2018/01/01   #クロール開始日(yyyy/mm/dd)
day       = 5            #クロール日数
printout  = off          #印刷の有無(on/off)
```

## work_data.xlsxについて
該当日の日付と勤務時間を取得します。
以下のような形式でクロールする日数分記録されます。

|日付|所定時間|深夜時間|深夜残業時間|
|:--:|:----:|:-----:|:--------:|
|2018/01/01|1|1|1|
|2018/01/01|1|1|1|
|...|...|...|...|
