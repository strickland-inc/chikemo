# Chikemo

## Project Map

- `メイン.gs`: Chikemoの既存スプレッドシート連携GAS。
- `Chikemo購入フォーム.gs`: Chikemo購入フォーム専用GAS。
- `corp-site/`: Chikemoコーポレートサイトの独立Gitリポジトリ。
- `_archive/`: 過去実装。明示依頼なしに復活・編集しない。
- `_logs/`: 運用ログ。

## Account Boundary

- Chikemo購入フォームのGoogle Workspace owner and execution account: `chikemo.info@chikemo.net`
- Expected Resend sender: `chikemo.info@chikemo.net`
- `メイン.gs`の顧客向け連絡先: `chikemo.info@gmail.com`
- Googleアカウント、送信元、reply-toを混同しない。
- CLIからGmailを送信しない。メール送信はApps Scriptまたは本番経路に限定する。

## Apps Script Projects

### Existing Chikemo GAS

- Apps Script project ID: `17GFyH04GH6v7BHCHYvqVMjpwUSMkadlu6BC2zZEDaZDElrirsokIovF_`
- Apps Script: `https://script.google.com/home/projects/17GFyH04GH6v7BHCHYvqVMjpwUSMkadlu6BC2zZEDaZDElrirsokIovF_/edit`
- Local clasp config: `.clasp.json`
- The bound Spreadsheet URL is not recorded in this repository. Do not guess it.

### Chikemo Purchase Form

- Spreadsheet: `https://docs.google.com/spreadsheets/d/1Pf2GPmzRdf32QlhX-OutkTD_fdkWcDfjVhOluNZy_0Q/edit`
- Apps Script: `https://script.google.com/home/projects/1D8gGYjvfinrWlm3wGXVNLIewoEVFX3Hs06Z0PRRrzmgunBQAw5WKzIKU/edit`
- Apps Script project ID: `1D8gGYjvfinrWlm3wGXVNLIewoEVFX3Hs06Z0PRRrzmgunBQAw5WKzIKU`
- Local clasp config: `.clasp.chikemo-purchase.json`
- Local clasp ignore: `.claspignore.chikemo-purchase`

Never run `clasp push` without explicitly selecting and verifying the intended config. The default and purchase-form configs point to different production projects.

## Mail Delivery

- Use Gmail first and Resend only for Gmail quota errors.
- Resend must call the `/emails` API directly; do not add a domain-preflight request to the send path.
- `RESEND_API_KEY` belongs in Apps Script script properties. Never save its value in files or Git.
- Installable edit triggers must not use the simple-trigger function name `onEdit`.

## Existing GAS Constraints

- `メイン.gs`のデータ投入元はLINEエルメ。Google Formsではないため`onFormSubmit`へ変更しない。
- `handleEdit`は半角名の`シート1`だけを処理する。
- 日時は`Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')`で文字列化してからセルへ書く。
- 入金列への複数行貼り付けでは先頭行だけを処理し、残りへ警告を書く。通常運用では1行ずつ入力する。
- R列とV〜AA列の自動転記はSpreadsheetのARRAYFORMULAで行い、`setupArrayFormulas()`で設定する。
- AJ列の処理監視はSpreadsheet関数で行い、`setupMonitoringFormula()`で設定する。監視列自体はメールを送信しない。
- 未送信の復旧には`reprocessUnsent()`を使用する（入金=OKの発送通知と入金=NGのキャンセル通知の両方が対象）。実行前に対象行と既送信ガードを確認する。
- `handleEdit`は`LockService`で同時編集を直列化し、入金列が編集範囲に含まれていれば処理する（先頭列ではなく包含で判定）。OK/NGは全角半角・大文字小文字・前後空白を無視して判定する。
- 自動送信されなかった通知は運用者が手動で送る。自動の補完送信（巡回トリガー）は入れない。手動で送った行はAD列/AG列に「手動送信済み」と入力し、再送とAJ列の警告の対象外にする。
- `removeTrigger()`は`handleEdit`のトリガーだけを消す。

## Git And Verification

- Git remote: `https://github.com/kochan17/chikemo.git`
- Before completion run `node --test tests/chikemoPurchaseForm.test.mjs tests/chikemoMain.test.mjs`, the relevant clasp status command, and `git diff --check`.
- Preserve unrelated work and do not commit `.clasprc.json`, OAuth tokens, cookies, API keys, or customer data.
