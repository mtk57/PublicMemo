# VB6版 xTextコントロールテストの使用ガイド

このガイドでは、VB6でコムラッド社のxTextコントロールをテストするための手順と注意点を説明します。

## 前提条件

- Visual Basic 6.0
- コムラッド社のFormDesignerコンポーネント（COMRADD.xText）がインストール済み
- Microsoft Windows Common Controls 6.0（SP6）（MSComctlLib.ListView）

## プロジェクトのセットアップ

### 1. 新規プロジェクトの作成

1. Visual Basic 6.0を起動
2. 「スタンダードEXE」プロジェクトを選択して新規作成

### 2. 必要なコンポーネントの参照設定

1. 「プロジェクト」メニュー→「コンポーネント」を選択
2. 以下のコンポーネントにチェックを入れる：
   - Microsoft Windows Common Controls 6.0（SP6）
   - コムラッド社 FormDesigner コンポーネント（COMRADD ActiveX Control）

### 3. テストフォームの作成

1. 提供されたコードを新しいフォームモジュールにコピー＆ペースト
2. フォームデザイナでコントロールを配置：
   - ListView コントロール
   - TextBox コントロール（テスト表示用）
   - CommandButton コントロール x 2
   - Label コントロール

## テストプログラムの使用方法

### テストの実行

1. プロジェクトを実行（F5キー）
2. テストフォームが表示されます
3. 「すべてのテストを実行」ボタンをクリックして一連のテストを開始
4. テスト結果がリストビューに表示されます

### テスト結果の見方

- **テスト名**: 実行されたテストの名前
- **結果**: 「成功」または「失敗」
  - 成功（緑色）: テストがパスした
  - 失敗（赤色）: テストが失敗した
- **詳細**: テストの説明またはエラーメッセージ

画面上部に合計のテスト統計（合計、成功、失敗の数）も表示されます。

## テストの種類

テストプログラムは以下の主要カテゴリのテストを実行します：

### 1. プロパティテスト

- **MaxLengthB_DefaultValueIsZero**: MaxLengthBプロパティの初期値が0かを確認
- **MaxLengthB_SetAndGetValue**: MaxLengthBプロパティが正しく設定・取得できるかを確認
- **MaxLengthB_NegativeValue_ThrowsError**: 負の値を設定するとエラーになるかを確認

### 2. テキスト入力制限テスト

- **Text_WithAsciiOnly_RespectsMaxLengthB**: 半角文字の入力がバイト数制限を尊重するか
- **Text_WithJapaneseOnly_RespectsMaxLengthB**: 全角文字の入力がバイト数制限を尊重するか
- **Text_WithMixedChars_RespectsMaxLengthB**: 半角・全角混在文字列がバイト数制限を尊重するか
- **MaxLength_And_MaxLengthB_SmallerValueApplied**: MaxLengthとMaxLengthBの小さい方が適用されるか

### 3. クリップボード操作テスト

- **Paste_RespectsMaxLengthB**: ペースト操作時にバイト数制限が適用されるか
- **Paste_WithSelection_ReplacesSelectedText**: 選択範囲がペーストで置き換えられるか
- **Paste_WithSelection_ExceedingMaxLengthB_Truncates**: 制限超過時に切り詰められるか

### 4. 特殊文字テスト

- **Text_WithSpecialJapaneseChars_RespectsMaxLengthB**: 漢字や特殊記号がバイト数制限を尊重するか

## VB.NETとの比較方法

VB6とVB.NETで同等のテストを実行し、結果を比較することで、移植したコントロールが元のコントロールと同じ動作をするかを確認できます。

### 比較のポイント

1. **バイト数計算**: 全角・半角文字のバイト数計算が同じか
2. **制限適用ロジック**: MaxLengthとMaxLengthBの適用優先順位が同じか
3. **ペースト動作**: クリップボード操作時の制限適用が同じか
4. **切り詰めロジック**: 制限超過時の文字列切り詰め動作が同じか

### 結果の記録

テスト結果を以下のような表形式で記録すると比較しやすくなります：

| テスト名 | VB6結果 | VB.NET結果 | 一致 |
|---------|--------|-----------|------|
| テスト1  | 成功    | 成功       | ✓   |
| テスト2  | 成功    | 失敗       | ✗   |

## 注意事項

1. **SendKeys関連**:
   - VB6でのクリップボードテストはSendKeysを使用しています
   - フォーカスやウィンドウのアクティブ状態によって動作が変わる場合があります
   - テスト実行中はフォームが最前面にあることを確認してください

2. **クリップボード**:
   - テスト実行前にクリップボードの内容が保存され、テスト後に復元されます
   - 何らかの理由でクリップボード操作が失敗した場合、該当テストはスキップされます

3. **コントロール参照**:
   - `COMRADD.xText`への参照はコムラッド社のFormDesignerがインストールされている環境が必要です
   - 参照が見つからない場合は、OLEコンテナコントロールを使って代替することもできます

4. **文字コード**:
   - VB6とVB.NETでは文字列の内部処理方法が異なります
   - 特に日本語などの全角文字の扱いに注意が必要です
