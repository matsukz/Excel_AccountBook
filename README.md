# 仕組み
## 新規取引
新規取引ボタンを押すとマクロ`Enter`が実行されます。<br>
表示されるウィンドウにテキストを入力すると、ワークシート`出費明細`にテーブル`メインテーブル`の最後にデータを追加します。

|テーブル上の項目|対応入力フォーム|対応変数|変数の型|
|:---:|:---:|:---:|:---:|
|日付|日付を入力|NewDate|Date|
|支払先|支払先を入力|Client|String|
|内容|内容を入力|Contents|String|
|分類|分類を入力|Classification|String|
|金額|金額を入力|Amount|Long|

日付フォームが空白または`mm/dd`の書式以外の場合は処理がキャンセルされます。


# 参考
FOM出版　よくわかるMicrosoft Excel 2019/2016/2013 マクロ/VBA 

[EXCEL VBA ドロップダウンリスト・プルダウンリスト・コンボボックスの作成（リスト選択）](https://akira55.com/drop_down_list/)採録日：2022年12月16日

[VBA 例外処理のサンプル(Excel/Access)](https://itsakura.com/excel-vba-exception)  再録日：2022年12月16日

[【マクロVBA】カウントをCOUNT・COUNTIF・COUNTIFSで求める！複数条件にも対応](https://dokugakuexcel.com/%E3%80%90%E3%83%9E%E3%82%AF%E3%83%ADvba%E7%9F%A5%E8%AD%98-20%E3%80%91%E3%82%AB%E3%82%A6%E3%83%B3%E3%83%88%E3%82%92count-countif-countifs%E3%81%A7%E6%B1%82%E3%82%81%E3%82%8B/)  再録日：2022年12月16日

# おわりに
作成：**Matsukz**<br>
最終更新日：**2022年12月16日**
