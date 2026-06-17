# 資産管理台帳アプリ（Excel .xlsm）実装ガイド

このディレクトリには、`asset_register_requirements.md` に基づく初期実装（VBAモジュール）を配置しています。

## 構成
- `src/modConstants.bas` : シート名・定数
- `src/modMasters.bas` : マスタ読み取り
- `src/modNumbering.bas` : 資産番号採番
- `src/modValidation.bas` : 入力検証
- `src/modAudit.bas` : 変更履歴記録（Windows ID含む）
- `src/modPrint.bas` : 単票印刷
- `src/modAssetEntry.bas` : 新規登録・更新のユースケース
- `master/*.csv` : 初期マスタデータ

## Excel側セットアップ手順
1. 新規ブックを `AssetRegister.xlsm` として保存。
2. 以下のシートを作成。
   - 台帳
   - マスタ_設置場所
   - マスタ_種別
   - マスタ_償却方法
   - 採番管理
   - 変更履歴
   - 印刷
   - 設定
3. `master` 配下CSVの内容を各マスタシートへ貼り付け。
4. VBAエディタで `src/*.bas` をインポート。
5. `modAssetEntry.InitializeWorkbook` を1回実行（ヘッダ生成）。

## 台帳列定義（A:Q）
A=資産番号 / B=資産種別コード / C=購入orリース / D=設置場所コード / E=名称 / F=型式 / G=取得年月日 / H=取得額 / I=耐用年数 / J=償却方法 / K=償却率 / L=管理番号 / M=備考 / N=状態 / O=最終更新日時 / P=最終更新者 / Q=最終更新WindowsID

## 印刷
`modPrint.PrintByAssetNumber` を呼び出し、対象資産番号を渡してください。
複数レコード印刷は仕様上サポートしません。
