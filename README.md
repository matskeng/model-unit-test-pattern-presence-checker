# model-unit-test-pattern-presence-checker

モデル単体テスト用 Excel ファイルにおける **単体テストパターンの有無** を確認するツールです。  
指定したフォルダ配下を再帰的に走査し、
`単体テストファイル_*.xls*` の内容から単体テストパターンの有無を判定し、
結果を **CSV ファイル**として出力します。

---

## 特徴

- 指定フォルダ配下を再帰的に走査
- `単体テストファイル_*.xlsx / *.xlsm` を自動検出
- Excel シート・セルを基にした機械的な有無判定
- 結果を **Excel でそのまま開ける CSV（UTF-8 BOM付き）** で出力

---

## 判定仕様

- 対象シート名：`テストパターン`
- 対象セル：`AP5`

| AP5 セルの状態 | 判定結果 |
|---|---|
| 値が入っている | 単体テストパターンあり |
| 空 | 単体テストパターンなし |

---

## インストール

```bash
pip install -r requirements.txt
```

---

## 使い方

```bash
python src/check_test_pattern_presence.py <走査対象フォルダ>
```

出力ファイル名を指定する場合：

```bash
python src/check_test_pattern_presence.py <走査対象フォルダ> -o result.csv
```

---

## 出力結果

- デフォルト出力ファイル：`pattern_check.csv`
- 文字コード：UTF-8（BOM付き）

出力項目：
- ファイルパス
- ファイル名
- 単体テストパターン有無（あり / なし / シートなし / 読込失敗）

---

## ライセンス

MIT License
