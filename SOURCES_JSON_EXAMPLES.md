# 取得元アプリ設定JSON 例

`取得元アプリ設定JSON` は、どのkintoneアプリから、どの条件で、どの列を、Excelのどのシートへ書き込むかを指定する設定です。

テンプレート管理アプリの `sources_json` に入力すると、その帳票専用の設定として使われます。

プラグイン設定画面のJSONは、テンプレート管理レコード側が未入力の場合の共通既定値です。

## 最小例

```json
[
  {
    "key": "actual",
    "label": "実績",
    "appId": "99",
    "sheetName": "日別計画実績",
    "fields": [
      { "code": "日付", "label": "日付", "type": "date" },
      { "code": "店舗名", "label": "店舗名" },
      { "code": "実績_総売上", "label": "実績_総売上", "type": "number" }
    ],
    "filters": [
      { "field": "店舗名", "operator": "=", "valueFrom": "store", "valueType": "text" },
      { "field": "日付", "operator": "between", "valueFrom": "dateRange", "valueType": "text", "dateRule": "monthStartToBaseDate" }
    ],
    "sorts": [
      { "field": "日付", "order": "asc" }
    ]
  }
]
```

## 複数アプリ例

```json
[
  {
    "key": "annual_budget",
    "label": "年初予算",
    "appId": "101",
    "sheetName": "年初予算",
    "fields": [
      { "code": "店舗名", "label": "店舗名" },
      { "code": "日付", "label": "日付", "type": "date" },
      { "code": "予算金額", "label": "予算金額", "type": "number" }
    ],
    "filters": [
      { "field": "店舗名", "operator": "=", "valueFrom": "store", "valueType": "text" },
      { "field": "日付", "operator": "between", "valueFrom": "dateRange", "valueType": "text", "dateRule": "yearStartToBaseDate" }
    ],
    "sorts": [
      { "field": "日付", "order": "asc" }
    ]
  },
  {
    "key": "actual",
    "label": "実績",
    "appId": "102",
    "sheetName": "日別計画実績",
    "fields": [
      { "code": "日付", "label": "日付", "type": "date" },
      { "code": "店舗名", "label": "店舗名" },
      { "code": "実績_総売上", "label": "実績_総売上", "type": "number" }
    ],
    "filters": [
      { "field": "店舗名", "operator": "=", "valueFrom": "store", "valueType": "text" },
      { "field": "日付", "operator": "between", "valueFrom": "dateRange", "valueType": "text", "dateRule": "monthStartToBaseDate" }
    ],
    "sorts": [
      { "field": "日付", "order": "asc" }
    ]
  }
]
```

## 既存Power Query由来の実績例

```json
[
  {
    "key": "actual",
    "label": "実績",
    "appId": "99",
    "sheetName": "日別計画実績",
    "fields": [
      { "code": "日付", "label": "日付", "type": "date" },
      { "code": "店舗名", "label": "店舗名" },
      { "code": "営業部", "label": "営業部" },
      { "code": "種類", "label": "種類" },
      { "code": "レート", "label": "レート", "type": "number" },
      { "code": "型", "label": "型" },
      { "code": "実績_営業日数", "label": "実績_営業日数", "type": "number" },
      { "code": "実績_台数", "label": "実績_台数", "type": "number" },
      { "code": "実績_総IN", "label": "実績_総IN", "type": "number" },
      { "code": "実績_総売上", "label": "実績_総売上", "type": "number" },
      { "code": "実績_総粗利", "label": "実績_総粗利", "type": "number" },
      { "code": "実績_総景品", "label": "実績_総景品", "type": "number" },
      { "code": "実績_利益率", "label": "実績_利益率", "type": "number" },
      { "code": "実績_再プレイ額", "label": "実績_再プレイ額", "type": "number" },
      { "code": "実績_補正総売上", "label": "実績_補正総売上", "type": "number" },
      { "code": "実績_補正総粗利", "label": "実績_補正総粗利", "type": "number" }
    ],
    "filters": [
      { "field": "店舗名", "operator": "=", "valueFrom": "store", "valueType": "text" },
      { "field": "日付", "operator": "between", "valueFrom": "dateRange", "valueType": "text", "dateRule": "monthStartToBaseDate" }
    ],
    "sorts": [
      { "field": "日付", "order": "asc" }
    ]
  }
]
```

## 日付フィールドがないアプリ例

店舗マスタなど、日付による絞り込みが不要なアプリでは対象期間の条件を設定しません。

```json
[
  {
    "key": "store_master",
    "label": "店舗マスタ",
    "appId": "201",
    "sheetName": "店舗マスタ",
    "fields": [
      { "code": "店舗ID", "label": "店舗ID" },
      { "code": "店舗名", "label": "店舗名" },
      { "code": "営業部", "label": "営業部" }
    ],
    "filters": [
      { "field": "店舗ID", "operator": "=", "valueFrom": "store", "valueType": "text" }
    ],
    "sorts": [
      { "field": "店舗ID", "order": "asc" }
    ]
  }
]
```

`filters` を空配列にすると、そのアプリは全件取得します。

## 項目説明

| 項目 | 必須 | 説明 |
|---|---:|---|
| `key` | 任意 | 内部識別名。英数字推奨 |
| `label` | 任意 | 人間向けの名前 |
| `appId` | 必須 | 取得元kintoneアプリID |
| `sheetName` | 必須 | Excelの書き込み先シート名 |
| `fields` | 必須 | 出力するフィールドと列順 |
| `fields[].code` | 必須 | kintoneのフィールドコード |
| `fields[].label` | 任意 | Excelの列見出し |
| `fields[].type` | 任意 | `text` / `number` / `date` / `datetime` / `boolean` |
| `filters` | 任意 | 複数指定できる抽出条件。空なら全件取得 |
| `filters[].field` | 必須 | 条件に使うフィールドコード |
| `filters[].operator` | 必須 | `=` / `!=` / `>` / `>=` / `<` / `<=` / `like` / `not like` / `in` / `not in` / `between` |
| `filters[].valueFrom` | 必須 | `store` / `dateRange` / `baseDate` / `fixed` |
| `filters[].value` | 条件付き | `fixed` の固定値。`in`系はカンマ区切り |
| `filters[].valueType` | 任意 | `text` または `number` |
| `filters[].dateRule` | 条件付き | `dateRange` の期間計算ルール |
| `sorts` | 任意 | 複数指定できる並び順 |
| `sorts[].field` | 必須 | 並び順に使うフィールドコード |
| `sorts[].order` | 必須 | `asc` または `desc` |

## 日付条件の期間ルール

| 値 | 意味 |
|---|---|
| `sameDay` | 基準日と同じ日 |
| `monthStartToBaseDate` | 基準日の月初から基準日 |
| `yearStartToBaseDate` | 1月1日から基準日 |
| `baseMonth` | 基準日の月初から月末 |
| `previousMonth` | 基準日の前月 |
| `nextMonth` | 基準日の翌月 |
| `baseMonthToNextMonthEnd` | 基準日の月初から翌月末 |

## 注意

- `fields` の順番がExcel列順になります。
- `fields[].code` はkintoneのフィールドコードです。フィールド名ではありません。
- `fields[].label` はExcelに出す見出しです。
- 数値として扱う列は `type: "number"` を指定してください。
- 日付として扱う列は `type: "date"` を指定してください。
- 複数の `filters` はANDで結合します。
- `sorts` は必要な場合だけ指定します。システム側で勝手にレコード順を変更しません。
