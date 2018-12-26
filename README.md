
# xlsx2redmine


## 説明

ruby から .xlsx 読んで Redmine へ REST 使って登録します。


## バージョン

> redmine: 3.4.3  
> ruby: 2.4.4  
> Excel: .xlsx が使えるバージョン  


## 事前準備

+ Redmine の REST API を有効化しといてください  

+ REST 利用するユーザーのAPIアクセスキーでソースの @api_key を書きかえてください  
	（ついでに @url も Redmine のURLへ）

+ bundle
  ```
  $ bundle install --path vendor/bundle
  ```


## 実行
```
$ bundle exec ruby run.rb
```


## 注意事項

+ 実行前にExcelブックは閉じてください。開いたままだと問題が起きる可能性があります  
	※ Redmine で振られた ID 更新を行うため


## ライセンス

MIT
