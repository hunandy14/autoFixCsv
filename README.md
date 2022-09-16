自動修復CSV檔案中的括號
===

使用方法
```ps1
# 快速使用
irm bit.ly/autoFixCsv|iex; autoFixCsv "sample.csv"
```

詳細用法
```ps1
# 載入函式庫
irm bit.ly/autoFixCsv|iex;
# 轉換並自動生成 sample1_fix.csv
autoFixCsv 'sample1.csv'
# 轉換並自動生成 sample1_fix.csv 且消除所有項目的前後空白
autoFixCsv 'sample1.csv' -TrimValue
# 轉換到 sample1_fix.csv
autoFixCsv 'sample1.csv' 'sample1_fix.csv'
```
