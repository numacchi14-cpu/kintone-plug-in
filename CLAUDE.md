# Claude Code 引継ぎ: Kintone Excel帳票出力プラグイン

## 目的と仕様書

このリポジトリは、Kintoneから取得したデータを既存書式のExcel帳票へ流し込むプラグインである。業務要件と受入基準は、親プロジェクト `C:\Projects\PL Management\SPEC.md` を正とする。プラグイン固有の設計・互換性・検証結果は [SPEC.md](SPEC.md) を正とする。

PL Management案件を引き継ぐ場合は、最初に親プロジェクトの `CLAUDE.md` と本ファイル・`SPEC.md` を確認する。実運用の接続先、APIトークン、パスワード、アプリIDはコミットしない。

## 実装範囲

- テンプレート管理アプリで、初期テンプレート作成、完成版テンプレート添付、テンプレート検証を行う。
- 帳票出力アプリで、指定レコードの基準日または入力基準日からKintoneを絞り込み、Excel帳票をダウンロードする。
- Kintone APIは `kintone.api()` とカーソルAPIを使い、500件単位で取得する。
- テンプレートの `__PL_FORMULA__:=...` は、出力時にExcelの通常数式へ変換する。

## PL Managementの現行状態（2026-08-12）

- 現行リリース: **v0.16.4**
- 最新コミット: `9bc17b4 Format exported date as Excel date`
- 配布ZIP: `dist/kintone-excel-report-plugin-v0.16.4.zip`
- 管理アプリ: 179、出力アプリ: 180。ただし環境固有値としてソースや設定例へ固定しない。
- 帳票ID: `monthly_department_pl`（単月）、累計部門別損益計算書もテンプレート・取得元設定を作成済み。

### v0.16.2～v0.16.4で直したこと

- **v0.16.2**: 数式プレースホルダーを変換するときに計算結果`0`をキャッシュとして保存しない。出力ブックの `calcPr` に `calcMode=auto`、`fullCalcOnLoad=1`、`forceFullCalc=1` を設定し、Excel起動時に全再計算する。
- **v0.16.3**: Kintoneから来る `yyyy-mm-dd` のISO文字列をExcel日付に変換する。設定JSONが旧来の`text`型でも日付条件付き `SUMIFS` が0件にならないようにした。
- **v0.16.4**: `exportedAt` もExcel日付として設定シートのB7へ書き込む。帳票の `TEXT(設定!B7,"[$-ja-JP-x-gannen]ggge年m月d日")` が和暦を表示できる。

## PL帳票の検証状況

- 単月（2025年11月）は旧帳票との1,518セル全件比較が差異0件。
- 累計（2025年1～11月）はテンプレート `C:\Projects\PL Management\outputs\cumulative-template-20260812\累計部門別PL_Kintoneテンプレート_完成版_千円表示_修正版8.xlsx` と取得元JSONを作成済み。
- 累計は実Kintone出力で金額・書式を目視確認済みで、出力日もv0.16.4で正常化済み。ただし旧累計帳票との沖縄・九州の全セル照合は未完了。
- 累計の元データ範囲は `INDEX` / `COUNTA` で動的に取得する。固定1万行参照、空のExcelテーブル、旧帳票の値の固定コピーは使わない。

## 主なソース

- `src/desktop.ts`: レコード詳細画面・一覧画面の出力UI
- `src/config.ts`: プラグイン設定画面
- `src/shared/excel.ts`: Excel読込・取得データの書込・数式変換・自動再計算指定
- `src/shared/kintoneApi.ts`: KintoneクエリとカーソルAPI
- `src/shared/dateRules.ts`: 基準日・期間条件
- `src/shared/config.ts`: 設定JSONの解析・保存
- `src/shared/types.ts`: 共通型
- `scripts/build.mjs`: TypeScript/Viteビルド
- `scripts/build-plugin.mjs`: プラグインZIP作成
- `plugin/manifest.json`: マニフェストとバージョン

## 開発・検証コマンド

PowerShellでは `npm` ではなく `npm.cmd` を使う。

```powershell
npm.cmd run typecheck
npm.cmd test
npm.cmd run package
```

リリース前に必ず型チェック・テスト・パッケージ作成を実行する。パッケージは `dist/kintone-excel-report-plugin-v<version>.zip` に生成される。

## 変更時のルール

- バグ修正・仕様変更時は、`package.json`、`package-lock.json`、`plugin/manifest.json` のバージョンを揃える。
- `SPEC.md` と本ファイルへ、変更理由、影響する帳票、確認方法、残作業を同じ変更単位で記録する。
- 生成済みZIPは動作確認用であり、秘密情報を含めない。
- 日付はExcel上で日付型として扱う。文字列の日付を式の比較・集計条件に渡さない。
- 出力テンプレートの見た目・セル結合・数値書式を壊さない。ExcelJSで読めることと、Excelで警告なしに開けることを確認する。
- Kintoneのドロップダウン条件は `=` ではなく `in` を使用する。`aggregation_status` を `=` で検索すると `GAIA_IQ03` になる。
- 金額は円で取得・計算し、帳票の書式でのみ千円表示・`△`表示にする。

## 次の作業

0. **作業前に`git push`でoriginへ退避する。** 現在ローカルの`agent/documentation-status`ブランチはoriginへ9コミット未プッシュ、かつ`main`へも未マージ（2026-08-12精査で確認）。直近のv0.16.x一連の変更が作業用PC1台にしか存在しない状態のため、区切りの良いタイミングで`git push`し、`main`へのマージも検討する。詳細は`SPEC.md`の「Claude Codeによる精査（2026-08-12）」を参照。
1. **単月の日付不具合をv0.16.5で修正する。** `src/shared/excel.ts` のISO日付変換がローカル日時を出力し、単月の完全一致`SUMIFS`が一致しない。UTC基準の時刻なし日付へ変更する。
2. 単月テンプレートの右上出力日式を日付型対応へ1セル変更する。累計テンプレートは変更しない。
3. 2025年11月単月で金額が再表示されること、同じ条件の累計出力で金額が変わらないことを確認する。
4. その後に累計の旧帳票との全セル比較を行う。詳細な原因と検証条件は`SPEC.md`の「単月の日付不具合と次の修正」を参照する。
