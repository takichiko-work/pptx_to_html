# pptx_to_html 処理フローチャート

## メイン処理（簡潔版）

```mermaid
flowchart TD
    A[PowerPointファイル読み込み] --> B[スライド要素抽出]
    B --> C[座標でソート]
    C --> D[GPT API呼び出し]
    D --> E[HTML断片生成]
    E --> F[テンプレート適用]
    F --> G[HTMLファイル出力]
    G --> H[完了]
```

## 詳細処理

### 1. PowerPointファイル読み込み
```mermaid
flowchart TD
    A[main.py開始] --> B[inputフォルダ検索]
    B --> C{PowerPointファイル存在?}
    C -->|No| D[エラー: ファイルなし]
    C -->|Yes| E[Presentationオブジェクト作成]
    E --> F[スライド範囲設定]
    F --> G[スライド幅・高さ取得]
```

### 2. スライド要素抽出
```mermaid
flowchart TD
    A[スライド要素抽出開始] --> B[全シェイプをループ]
    B --> C{シェイプタイプ判定}
    C -->|PICTURE| D[画像として処理]
    C -->|TEXT_FRAME| E[テキスト抽出]
    C -->|GROUP| F[子シェイプ再帰処理]
    C -->|TABLE| G[テーブルセル処理]
    C -->|その他| H[未対応シェイプ]
    D --> I[座標情報付きでリスト追加]
    E --> J{画像プレースホルダー?}
    J -->|Yes| K[画像として処理]
    J -->|No| L[テキストとして処理]
    K --> I
    L --> I
    F --> I
    G --> I
    H --> I
    I --> M{全シェイプ処理完了?}
    M -->|No| B
    M -->|Yes| N[要素抽出完了]
```

### 3. 座標でソート
```mermaid
flowchart TD
    A[座標ソート開始] --> B[座標正規化]
    B --> C[top座標でソート]
    C --> D[left座標でソート]
    D --> E[不要要素除外]
    E --> F[ヘッダー・フッター除外]
    F --> G[パンくずリスト除外]
    G --> H[ソート完了]
```

### 4. GPT API呼び出し
```mermaid
flowchart TD
    A[GPT処理開始] --> B[ルールファイル読み込み]
    B --> C[パーツ名自動判定]
    C --> D[パーツルール取得]
    D --> E[プロンプト構築]
    E --> F[API入力ログ出力]
    F --> G[OpenAI API呼び出し]
    G --> H{API呼び出し成功?}
    H -->|No| I[エラー処理]
    H -->|Yes| J[API出力ログ保存]
    J --> K[HTML断片取得]
    I --> L[エラー終了]
    K --> M[GPT処理完了]
```

### 5. HTML断片生成
```mermaid
flowchart TD
    A[HTML生成開始] --> B[GPT応答解析]
    B --> C[HTMLタグ生成]
    C --> D[クラス名適用]
    D --> E[画像タグ生成]
    E --> F[テーブル構造生成]
    F --> G[HTML断片完成]
```

### 6. テンプレート適用
```mermaid
flowchart TD
    A[テンプレート適用開始] --> B[template.html読み込み]
    B --> C[{contents}置換]
    C --> D[{pagettl}置換]
    D --> E[HTML構造完成]
```

### 7. HTMLファイル出力
```mermaid
flowchart TD
    A[ファイル出力開始] --> B[outputディレクトリ確認]
    B --> C[HTMLファイル作成]
    C --> D[エンコーディング設定]
    D --> E[ファイル書き込み]
    E --> F[出力完了]
```
