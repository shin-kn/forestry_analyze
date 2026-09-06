# forestry

日本の林業に関する調査・資料作成。**このプロジェクトでは日本語で対応する。**

## 公開に関するルール（最重要）

このリポジトリは**共有用**。公開してよいのは次の2種類だけ。

1. **ネットから誰でも取れる資料**（林野庁の白書、各国政府の公開レポート、公開統計など）
2. **1から推論・分析して作成した資料**（比較レポート、集計、考察など）

**それ以外を公開しそうになったら、作業を止めてユーザーに確認すること。**

止める対象の例:
- `ref/internship_program/` の資料（インターンシップ・講演資料。`.gitignore` 済み）
- `user_docs/`、`ref_for_user/` に置かれたもの
- 出所が不明な資料、ユーザーから個別に受け取った資料

「公開」には以下を含む:
- git へのコミット・プッシュ
- Artifact としての publish
- 外部サービスへの送信

判断に迷ったら公開せず、確認する。

## リポジトリ運用

- ブランチは切らない。共有目的の単一ブランチで運用する。
- `.gitignore` は Python 関連、`*.pdf`、`ref/internship_program/` を除外。

## ディレクトリ構成

| パス | 内容 |
|---|---|
| `ref/official/white_paper/japan/md/` | 令和7年度 森林・林業白書を Markdown 化したもの（82ファイル） |
| `ref/official/white_paper/austria/` | オーストリアの森林・林業関連PDF（gitignore対象） |
| `ref/internship_program/` | インターンシップ資料（**非公開**、gitignore対象） |
| `quantitative/overview.md` | 白書から抽出した定量データ集 |
| `reports/` | 分析レポート |
| `pptx/` | スライド生成スクリプトと成果物 |

## 環境

- Python は `uv.exe` 経由で実行する（WSL側に Python はない）
- `cd` は使わず絶対パスを使う
