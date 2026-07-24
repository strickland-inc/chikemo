# Chikemo購入フォーム GAS

`Chikemo購入フォーム.gs` は、次の本番Apps ScriptをGit管理します。

- Apps Script: `1D8gGYjvfinrWlm3wGXVNLIewoEVFX3Hs06Z0PRRrzmgunBQAw5WKzIKU`
- Spreadsheet: `1Pf2GPmzRdf32QlhX-OutkTD_fdkWcDfjVhOluNZy_0Q`
- 実行アカウント: `chikemo.info@chikemo.net`

既存の `.clasp.json` は別のChikemoプロジェクトを参照するため変更しません。購入フォーム専用GASを操作する場合は、明示的に次の設定を使用します。

```bash
clasp -P .clasp.chikemo-purchase.json -I .claspignore.chikemo-purchase status
```

`RESEND_API_KEY` はApps Scriptのスクリプトプロパティで管理し、Gitへ保存しません。`.clasprc.json`、OAuthトークン、顧客データもコミットしません。
