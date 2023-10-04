# iso_audit_data_export

## プロジェクトの概要

このスクリプトは、ISO 監査用の必要な固定ファイルを生成します。

## プロジェクトの目的

プロジェクトの目的は、ISO 監査のために必要な文書を生成し、整理することです。

## 動作環境

```iso_audit_data_export.xlsm```はWindowsプラットフォーム専用です。他のオペレーティングシステム（例：macOS、Linux）では正しく動作しないことにご注意ください。

## 使用方法

1. このリポジトリをクローンまたはダウンロードして、ローカルマシンに保存します。

1. ```programs/vba/iso_audit_data_export.xlsm``` を開きます。

1. Excel ウィンドウ上部の「開発」タブをクリックし、そこから「Visual Basic」アイコンをクリックして VBA エディタを開いてください。  
   ```ConstantsModule``` モジュールの ```debugFlag``` を設定してください。このフラグで出力先フォルダの制御を行います。  
   出力先フォルダの設定は、```ClassEditPath``` クラスの ```GetOutputPath``` メソッドで行います。必要に応じて修正してください。

1. 実行ボタンをクリックしてください。処理が終了すると、その旨のポップアップメッセージが出力されます。

## 参照設定

このプロジェクトでは、```Microsoft Word 16.0 Object Library``` および ```Microsoft Scripting runtime``` の 2 つの外部ライブラリへの参照が必要です。

## programs/vba/modules 配下のプログラムの概要

### ClassEditPath

- 入出力フォルダパスを操作するためのメソッドとプロパティを提供します。

### ClassFolderPathManager

- フォルダ名のリストと処理対象のファイル名リストを操作するためのメソッドとプロパティを提供します。

### ConstantsModule

- Excel ファイル内の定数とメインの処理を提供します。

### CreateText

- テキストファイルを生成するための機能を提供します。

### ExportVba

- ExportVbaFiles()を実行すると、Thisworkbook 内のすべての VBA モジュールが指定されたディレクトリにエクスポートされ、外部ファイルとして保存されます。

### FileUtils

- ファイルを操作するためのユーティリティ関数を含むモジュールです。

### Utils

- その他のユーティリティ関数を含むモジュールです。

### convertToPdf

- Excel ファイルを PDF に変換するための機能を提供します。

## programs/R 配下のプログラムの概要

### for_test.R

- R 言語でのテスト用プログラムです。tidyverse パッケージがインストールされていない場合は、実行前にインストールしてください。

## tools 配下のプログラムの概要

### iconv.sh

- ファイルの文字コードを変換するためのシェルスクリプトです。
