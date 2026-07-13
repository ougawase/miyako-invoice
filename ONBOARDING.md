# はじめての開発ガイド（新メンバー向け）

対象：このリポジトリで初めて開発する人。GitHubでの共同開発が初めてでも、
上から順にやれば「最初のPull Request（変更の提案）」まで出せます。
所要：環境づくり30〜60分＋最初の課題。

> 先に言葉を3つだけ。
> **ブランチ** ＝ コードの「下書きコピー」。main（本番の正本）を直接触らず、自分専用の下書きを作って編集する。
> **Pull Request（PR）** ＝「この下書きを本番に取り込んでください」という提案。
> **マージ** ＝ 提案が承認されて本番に取り込まれること（このリポジトリでは平戸がやる）。

## 0. 最初の1回だけやる準備

### アカウントと道具

1. **GitHubアカウント**を作る（https://github.com ）→ ユーザー名を平戸に伝えて招待してもらう → 届いたメールの「Accept invitation」を押す
2. **Git** を入れる
   - Mac：ターミナル（Cmd+Space →「ターミナル」と検索して起動）で `xcode-select --install`
     （「already installed」というエラーが出たら、もう入っているのでそのまま次へ）
   - Windows：https://git-scm.com から「Git for Windows」を入れる（以後のコマンドは付属の「Git Bash」で実行）
3. **GitHub CLI（gh）** を入れる（GitHubへのログインが一番簡単になる道具）
   - Mac：ターミナルで `brew install gh`。`brew: command not found` と出たら、先に https://brew.sh の1行コマンドでHomebrewを入れる（5分）
   - Windows：https://cli.github.com の「Download for Windows」のインストーラでOK。
     インストール後、開いているGit Bashを一度閉じて開き直し、`gh --version` で確認
4. **VS Code**（エディタ）：https://code.visualstudio.com
5. 実行環境（このリポジトリで使う方だけでOK）
   - yorisoi-site：**Node.js LTS**（https://nodejs.org ）
   - miyako-invoice：**Python**。Macは手順2で一緒に入るもの（3.9）でそのままOK。
     Windowsは https://python.org から（インストール時に「**Add python.exe to PATH**」に必ずチェック）

### 自分の名前を設定してログイン

```bash
git config --global user.name "自分の名前"
git config --global user.email "GitHubに登録したメールアドレス"
gh auth login
```

`gh auth login` の質問には順に：**GitHub.com** → **HTTPS** → **Authenticate Git ...? は Y（そのままEnter）** → **Login with a web browser**。
8桁のコード（XXXX-XXXX）が表示されるのでメモ → Enterを押すと開くブラウザにそのコードを入力。
他に質問が出たら基本そのままEnterでOK。

> **Windowsの人だけ**：Git Bashだと `gh auth login` の質問画面が動かないことがあります（既知の問題）。
> その場合は `winpty gh auth login` と打つか、スタートメニューの「PowerShell」で1回だけ実行してください
> （ログインは共有されるので、以後はGit BashでOK）。

## 1. コードを手元に持ってくる（clone・最初の1回だけ）

```bash
gh repo clone ougawase/yorisoi-site      # ホームページの人はこちら
gh repo clone ougawase/miyako-invoice    # 請求書ツールの人はこちら
cd yorisoi-site        # または cd miyako-invoice
```

フォルダは「ホームフォルダ直下」にできます（例：/Users/自分の名前/yorisoi-site）。

## 2. 動かしてみる

README.md の手順どおりにセットアップして起動する。
**Windowsの人へ**：READMEの `python3` は `python`、`pip3` は `pip` と読み替えてください（例：`python -m streamlit run app.py`）。

**ここで動くところまで確認してから**次へ（動かないまま編集を始めない）。

## 3. 変更を出すまでの一本道（毎回この順番）

```bash
git checkout main
git pull                        # ① 最新を取り込む
git checkout -b fix/date-sort   # ② 作業用ブランチ（下書き）を切る。名前は半角英数とハイフンで内容がわかるように

# ③ VS Codeで編集する：このフォルダにいる状態で
#      code .
#    と打つとVS Codeで開く（code が無いと言われたら：VS Codeを起動 → Cmd/Ctrl+Shift+P →
#    「Shell Command: Install 'code' command in PATH」を実行。または「ファイル→フォルダーを開く」でcloneしたフォルダを選ぶ）

# ④ 動作確認（READMEのテスト・ビルドを全部グリーンに）

git status                      # ⑤ 何が変わったか自分の目で確認（赤=未保存の変更）
git add -A
git commit -m "何をしたか（日本語でOK）"
git push -u origin HEAD         # ⑥ 自分のブランチをGitHubに送る
gh pr create --fill             # ⑦ Pull Request を作る
```

⑦を実行するとPRのURLが表示されます。ブラウザで開いて、説明欄（鉛筆マークで編集）に
**何を・なぜ・どう確認したか（スクショ歓迎）** を書いてください。
あとは平戸がレビューして、OKならマージ（＝本番反映）します。

いま自分がどのブランチにいるかは、いつでも `git branch` で確認できます（`*` が現在地）。

## 4. レビューで「ここ直して」と言われたら

別の日に再開するときは、まず `git branch` で自分のブランチ名を確認 → `git checkout 自分のブランチ名` で戻る。
そのあと、編集 → `git add -A` → `git commit -m "指摘対応"` → `git push`。
これだけで**同じPRに自動で追加**されます（新しいPRを作り直さない）。

## 5. 絶対ルール（3つだけ）

1. **mainブランチに直接pushしない**（必ずブランチ→PR）。
   うっかりmainのまま編集してしまっても、その場で `git checkout -b fix/名前` と打てば変更ごと新しいブランチに移れるので大丈夫（消さなくていい）
2. `--force` のつくコマンドと `git reset --hard` は使わない（困ったら壊す前に相談）
3. 分からなくなったら、フォルダごと消してcloneし直してOK。
   **ただし消えないのは「GitHubにpush済みの分」だけ。** 消す前に必ず `git status` を見て、
   残したい変更があれば `git add -A && git commit -m "途中まで" && git push -u origin HEAD` で先にGitHubへ送ってから消すこと

## 6. 詰まったら

- エラーメッセージを**そのままコピペ**して、IssueのコメントかLINEへ（「何をしたら・何が出たか」をセットで。スクショ歓迎）
- ChatGPTやClaudeに聞くのも全然アリ。ただし出てきたコードを使う前に、このリポジトリの CONTRIBUTING.md のルールに合っているか確認すること
