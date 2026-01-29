# ScreenShot_Outlook

スクリーンショットを自動的に撮影し、Outlookメールに添付して送信準備を行うPowerShellツールです。キーボードショートカット（Ctrl+Alt+S）で素早く実行できます。

## 機能

- プライマリスクリーンのスクリーンショット自動撮影
- Outlookメール新規作成と画像添付
- 日付別フォルダに画像を自動保存
- キーボードショートカット（Ctrl+Alt+S）対応
- バックグラウンド実行（ウィンドウ非表示）
- 設定ファイルによる宛先・件名のカスタマイズ

## 必要環境

- Windows 10/11
- PowerShell 5.1以上
- Microsoft Outlook（インストール済み）

## インストール

1. このリポジトリをクローンまたはダウンロード：
```bash
git clone https://github.com/SHunsuke-NIshijima/ScreenShot_Outlook.git
cd ScreenShot_Outlook
```

2. 設定ファイルをカスタマイズ：
`config.json`または`config-new.json`を編集し、送信先メールアドレスや件名を設定します。

```json
{
  "name": "あなたの名前",
  "mailto": "送信先@example.com",
  "cc": "cc@example.com",
  "subject": "スクリーンショット"
}
```

3. ショートカットを登録：
```powershell
powershell -ExecutionPolicy Bypass -File Setup-Shortcut.ps1
```

これにより、Ctrl+Alt+Sでスクリーンショット送信が可能になります。

## ファイル構成

| ファイル名 | 説明 |
|-----------|------|
| `Send-Screenshot-Email-New.ps1` | メインスクリプト（最新版） |
| `Send-Screenshot-Email.ps1` | メインスクリプト（旧版） |
| `Run-Hidden.vbs` | PowerShellをバックグラウンドで実行するVBScript |
| `Setup-Shortcut.ps1` | キーボードショートカット設定スクリプト |
| `config-new.json` | 設定ファイル（最新版） |
| `config.json` | 設定ファイル（旧版） |
| `Fix-Encoding.ps1` | エンコーディング修正用ユーティリティ |

## 使い方

### 方法1: キーボードショートカット
インストール後、**Ctrl+Alt+S**を押すだけで実行されます。

### 方法2: 手動実行
```powershell
powershell -ExecutionPolicy Bypass -File Send-Screenshot-Email-New.ps1
```

### 方法3: バックグラウンド実行
```cmd
wscript Run-Hidden.vbs
```

## 動作フロー

1. スクリプト実行
2. プライマリスクリーンのスクリーンショット撮影
3. `image/YYYYMMDD/screenshot_HHMMSS.png`として保存
4. Outlookで新規メール作成
5. 撮影した画像を添付＆本文に埋め込み
6. メールウィンドウを表示（送信は手動）

## トラブルシューティング

### スクリプト実行エラー
実行ポリシーの制限がある場合：
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Outlookエラー
- Outlookが起動していることを確認
- Outlookのセキュリティ設定を確認

### 文字化け問題
`Fix-Encoding.ps1`を実行してファイルエンコーディングを修正：
```powershell
powershell -ExecutionPolicy Bypass -File Fix-Encoding.ps1
```

## カスタマイズ

### ショートカットキーの変更
`Setup-Shortcut.ps1`の以下の行を編集：
```powershell
$shortcut.Hotkey = "CTRL+ALT+S"  # 好きなキーに変更
```

### 保存先フォルダの変更
`Send-Screenshot-Email-New.ps1`の以下の行を編集：
```powershell
$imageFolder = Join-Path $scriptPath "image\$date"
```

## ライセンス

このプロジェクトはオープンソースです。自由に使用・改変できます。
