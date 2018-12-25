
# xlsx_to_redmine


## 説明

ruby から .xlsx 読んで Redmine へ REST 使って登録します。


## 対象バージョン
> redmine 3.4.3  
> ruby 2.4.4  
> Excel .xlsx が使えるバージョン  


## 事前準備

+ この辺りを参考に Redmine の REST API を有効化しといてください
http://blog.redmine.jp/articles/redmine-ticket-ikkatsu/

+ REST 利用するユーザーのAPIアクセスキーでソースの @api_key を書きかえてください
　(ついでに @url も Redmine のURLへ)

```
$ bundle install --path vendor/bundle
```


## 実行
```
$ bundle exec ruby rest.rb
```


## 注意事項

実行前に wbs.xlsx は閉じてください。開いたままだと問題が起きる可能性があります。
※ Redmine で振られた ID 更新を行うため


## ライセンス

MIT
