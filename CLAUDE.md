# CLAUDE.md

このファイルは、このリポジトリで作業する Claude Code 向けのガイドです。

## プロジェクト概要

Google Apps Script（GAS）のスクリプト集です。

## 開発環境

- ランタイム: Google Apps Script（V8 ランタイム）
- ローカル開発・デプロイには [clasp](https://github.com/google/clasp) を使用する想定
  - `clasp login` — Google アカウントで認証
  - `clasp clone <scriptId>` — 既存プロジェクトを取得
  - `clasp push` — ローカルの変更を Apps Script に反映
  - `clasp pull` — Apps Script 側の変更をローカルに取得
  - `clasp open` — ブラウザでスクリプトエディタを開く

## コーディング規約

- 言語は JavaScript（`.gs` / `.js`）。`const` / `let` を使い、`var` は使わない
- GAS の実行時間制限（1 回 6 分）を意識し、大量データはバッチ処理（`getValues()` / `setValues()` でまとめて読み書き）する
- API キーやスプレッドシート ID などの秘密情報・環境依存値はコードに直書きせず、`PropertiesService` のスクリプトプロパティで管理する
- トリガーで実行される関数は、関数名とトリガー種別をコメントで明記する

## Git 運用ルール

- **コードを変更するたびに、コミットして GitHub にプッシュすること。**
  - 1 つの変更（機能追加・修正など）が完了したら、その都度 `git add` → `git commit` → `git push` を行う
  - 複数の無関係な変更をまとめて 1 コミットにしない
- コミットメッセージは日本語で、変更内容が分かるように簡潔に書く
  - 例: `売上集計スクリプトに月次集計機能を追加`
- `.clasp.json`（スクリプト ID を含む）や認証情報（`.clasprc.json` など）は `.gitignore` に含め、コミットしない
- プッシュ前に `git status` / `git diff` で意図しないファイルが含まれていないか確認する
