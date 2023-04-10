
# AccdbExport

## はじめに

- このアプリはAccdbファイルからVBA、Form、テーブル定義、クエリ(SQL)の情報を抽出してテキスト形式のファイルに書き出します。
- 出力されたファイルをWinMergeなどで比較すれば、Accessアプリの変更内容を簡単に確認することができます。
- 元は [Accessで必ず使うVBA ほぼ完全なエクスポート 全オブジェクト+参照設定リストをエクスポート - Qiita](https://qiita.com/Q11Q/items/99caf05417f1c3af850c) で公開されているVBScriptを使用していたのですが、時間がかかる、不要な情報も抽出される、アバスト無料アンチウイルスがマルウェアとして認識する、といった不満があったので、C#で必要な機能だけを作成してみました。

## 使用方法

- TextBoxにAccdbファイルのパス名を入力してRunボタンをクリックすると、Accdbファイルのフォルダにフォルダが作成されて抽出したファイルが格納されます。
- TextBoxには、エクスプローラからAccdbファイルをマウスドラッグすることでパス名を入力することができます。
- 処理が正常に完了した場合は問題ないと思いますが、特にエラー終了した場合などにAccessのプロセスが残ることがあります。残ってしまったプロセスは、タスクマネージャから強制終了させてください。

## 注意

- アプリがデータを道連れにしてずっこける可能性がないとは言えません。そうなっても困らないように、Accdbファイルをコピーしてそのコピーをこのアプリに食わせるなどといった自衛をお願いします。

## 環境

- Microsoft Visual Studio Community 2022 (64ビット) Version 17.5.3
- WPF (.NET Framework 4.8)
- C# 7.3
- Microsoft® Access® for Microsoft 365 MSO (バージョン 2211 ビルド 16.0.15831.20098) 64 ビット
- Microsoft.Xaml.Behaviors.Wpf 1.1.39
- ReactiveProperty.Core 9.1.2

## 参考リンク

- [Accessで必ず使うVBA ほぼ完全なエクスポート 全オブジェクト+参照設定リストをエクスポート - Qiita](https://qiita.com/Q11Q/items/99caf05417f1c3af850c)
- [Visual Basic for Applicationsのアクセス オブジェクト モデル (VBA) | Microsoft Learn](https://learn.microsoft.com/ja-jp/office/vba/api/overview/access/object-model)
