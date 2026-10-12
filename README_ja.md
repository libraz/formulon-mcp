# formulon-mcp

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon-mcp/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon-mcp/actions)
[![npm](https://img.shields.io/npm/v/@libraz/formulon-mcp)](https://www.npmjs.com/package/@libraz/formulon-mcp)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon-mcp/blob/main/LICENSE)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net/ja/mcp/)

[Formulon](https://github.com/libraz/formulon) 用のローカル stdio MCP サーバーです。エージェントから Excel ワークブックの調査・編集、数式の再計算、印刷用帳票の作成、範囲のプレビュー、コピーした数式の監査、Excel テーブルとピボットテーブルの作成ができます。

**[詳細ドキュメント](https://formulon.libraz.net/ja/mcp/)** は Formulon のサイトにまとめています。ツールの入力・応答・制限と、具体的な作業例を掲載しています。

## クイックスタート

Node.js 22 以上が必要です。[`engines`](https://github.com/libraz/formulon-mcp/blob/main/package.json) も参照してください。同梱の対話式インストーラーでサーバーを登録します。

```sh
npx -y @libraz/formulon-mcp init
```

Claude Code、Codex CLI、Claude Desktop から対象を選び、続けて新しいセッションが計算に使う Excel のロケールとホスト（Windows または Mac）を選びます。ロケールは数値や日付の文字列、`TEXT` の書式、文字列の比較、2 バイト文字を扱う関数の結果を決めるもので、既定ではシステムのロケールが選ばれます。設定の変更内容を確認してクライアントを再起動します。エージェントに Formulon で `=SUM(10,20,30)` を計算するよう依頼すると、結果は `60` になります。

スクリプトから設定する場合は、`init --profile win-365-ja_JP` のようにプロファイルを直接指定します。個々のセッションでは、開くときに別のプロファイルを選べます。

手動設定、ローカルコードの利用、トラブルシュートは[インストール](https://formulon.libraz.net/ja/mcp/install)を参照してください。登録を削除するには `npx -y @libraz/formulon-mcp uninstall` を実行します。

## ガイド

| 作業 | ドキュメント |
| --- | --- |
| テンプレートを編集して別ファイルに保存する | [ワークブックのワークフロー](https://formulon.libraz.net/ja/mcp/workflow) |
| 数式と印刷設定を含む請求書を作る | [帳票の作成](https://formulon.libraz.net/ja/mcp/authoring) |
| 帳票を再現して見た目を確認する | [レイアウトとプレビュー](https://formulon.libraz.net/ja/mcp/layout-preview) |
| 数式の上書きや抜けを探す | [数式の監査](https://formulon.libraz.net/ja/mcp/formula-audit) |
| 売上データを集計する | [テーブルとピボット](https://formulon.libraz.net/ja/mcp/tables-pivots) |
| ツールの入力や低レベルメソッドを調べる | [ツール一覧](https://formulon.libraz.net/ja/mcp/tools)・[高度な API 操作](https://formulon.libraz.net/ja/mcp/advanced) |

セッションはメモリ上にあり、サーバー終了時に失われます。出力パスを明示し、読み込み・保存の診断を確認してください。プレビューは近似表示です。数式監査は確認候補を返します。[セキュリティモデル](https://formulon.libraz.net/ja/mcp/security)も参照してください。

## 開発

```sh
yarn install
yarn run check
yarn test
```

テストは `dist/` をビルドして実行します。計算エンジンは別の [Formulon リポジトリ](https://github.com/libraz/formulon)で管理しています。

## ライセンス

Apache-2.0。[LICENSE](https://github.com/libraz/formulon-mcp/blob/main/LICENSE) を参照してください。
