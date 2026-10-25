# CLAUDE.md

このファイルは、このリポジトリで作業する Claude Code へのガイドです。

## プロジェクト概要

Google Apps Script（GAS）のスクリプトを管理するプロジェクトです。

- 言語: JavaScript（GAS / V8 ランタイム）
- ローカル開発とデプロイには [clasp](https://github.com/google/clasp) を使う想定です

## 開発ルール

- GAS の V8 ランタイムで動く構文を使う（`const` / `let`、アロー関数、テンプレート文字列など）
- Node.js 専用の API（`require`、`fs` など）は GAS では使えないため使わない
- スプレッドシートの読み書きは `getValues()` / `setValues()` でまとめて行い、セル単位の API 呼び出しをループの中で繰り返さない（実行時間の上限 6 分に注意）
- API キーや ID などの秘密情報はコードに直接書かず、`PropertiesService` に保存する
- `.clasp.json` や `.clasprc.json` など認証情報を含むファイルはコミットしない（`.gitignore` に入れる）

## よく使うコマンド（clasp）

```bash
clasp login          # Google アカウントでログイン
clasp clone <scriptId>  # 既存の GAS プロジェクトを取得
clasp pull           # GAS 側の変更をローカルに取り込む
clasp push           # ローカルの変更を GAS にアップロード
clasp open           # ブラウザで GAS エディタを開く
```

## Git 運用ルール

**コードを変更したら、そのたびにコミットして GitHub にプッシュすること。**

1. 変更が終わったら `git status` と `git diff` で内容を確認する
2. 関係するファイルだけを `git add` する（`git add -A` で無関係なファイルや認証情報を含めない）
3. 何をなぜ変えたかがわかるコミットメッセージを書く
4. `git push` で GitHub にプッシュする
5. プッシュに失敗した場合（リモートとの競合など）は、強制プッシュせずにユーザーに報告する

補足:

- `git push --force` は、ユーザーが明示的に指示した場合を除いて使わない
- 1 回のコミットは 1 つのまとまった変更にする
- コミットメッセージは日本語でよい（例: `売上集計シートの日付フォーマットを修正`）
