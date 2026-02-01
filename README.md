# スクリーンショット自動メール作成ツール

## 概要
PowerShellとOutlookを使用して、スクリーンショットを撮影し、自動的に添付メールを作成するツールとなっている。ショートカットキー（例えば、Ctrl+Alt+S）を押すと、Windows標準のSnipping Tool（Win+Shift+S）のように範囲選択してスクリーンショットを撮影できる。

## 機能

### メイン機能
- **範囲選択スクリーンショット**: キーボードショートカットでマウスドラッグによる範囲選択キャプチャ（Win+Shift+S方式）
- **自動メール作成**: Outlookのメールに自動的に画像を添付し、件名・宛先を設定
- **画像の自動保存**: 日付別フォルダに整理して保存（形式: `image/yyyyMMdd/screenshot_HHmmss.png`）
- **直感的な操作**: 半透明オーバーレイと赤枠でリアルタイム表示

### 詳細な仕様

#### 1. スクリーンショット撮影
- **範囲選択方式**: マウスドラッグで任意の範囲を選択してキャプチャ
- **視覚的フィードバック**: 半透明黒オーバーレイ + 赤枠で選択範囲を表示
- **キャンセル機能**: ESCキーで中断可能
- PNG形式で保存
- ファイル名: `screenshot_HHmmss.png`（時分秒でユニーク）
- 保存先: `image/[日付]/` フォルダ（自動作成）

#### 2. メール自動作成
- **件名**: `【業務連絡】 （[名前]） yyyy/MM/dd`
- **宛先**: config.jsonで設定
- **CC**: config.jsonで設定
- **本文**: HTMLメール形式でスクリーンショットを埋め込み表示（幅1200px）
- **添付ファイル**: 撮影したスクリーンショット

#### 3. 設定管理
設定は `config.json` :
```json
{
  "name": "氏名",
  "mailto": "FirstName.LastName@gmail.com",
  "cc": "FirstName.LastName@gmail.com",
  "subject": "【業務連絡】{name}_{date}",
  "body": "【業務連絡】",
  "selectionBorderColor": "Red",
  "shortCutKey": "CTRL+ALT+S"
}
```

**設定項目**:

| 項目 | 説明 | 設定例 | デフォルト値 |
|------|------|--------|-------------|
| `name` | メール件名に表示される名前 | `"氏名"` | - |
| `mailto` | メールの宛先 | `"example@gmail.com"` | - |
| `cc` | メールのCC | `"cc@gmail.com"` | - |
| `subject` | メール件名のテンプレート<br>`{name}`: 名前に置換<br>`{date}`: 実行日（yyyy/MM/dd）に置換 | `"【業務連絡】 （{name}） {date}"` | `"【業務連絡】 （名前） 日付"` |
| `body` | メール本文テキスト<br>HTMLタグ使用可（例: `<br>`で改行） | `"スクリーンショットを添付します。"` | `"スクリーンショットを添付します。"` |
| `selectionBorderColor` | 範囲選択時の枠の色<br>.NET Color構造体の名前付き色 | `"Cyan"`, `"Red"`, `"Blue"`, `"Green"` | `"Red"` |
| `shortCutKey` | ショートカットキー<br>CTRL, SHIFT, ALTの組み合わせ | `"CTRL+ALT+S"`, `"CTRL+SHIFT+S"` | `"CTRL+ALT+S"` |

## ファイル構成

```
Outlook_PowerShell/
├── config.json                      # 設定ファイル（名前、メールアドレス）
├── Send-Screenshot-Email.ps1        # メインスクリプト
├── Run-Hidden.vbs                   # バックグラウンド実行用VBSスクリプト
├── Setup-Shortcut.ps1               # ショートカット登録スクリプト
├── Setup-Shortcut.bat               # ショートカット登録用バッチファイル（ダブルクリック実行可能）
├── image/                           # スクリーンショット保存フォルダ
│   ├── 20260129/
│   ├── 20260130/
│   └── 20260201/
└── README.md                        # 本ファイル
```

### 各ファイルの役割

| ファイル名 | 役割 |
|----------|------|
| `config.json` | ユーザー設定（名前、メールアドレス、CC、枠の色、ショートカットキー） |
| `Send-Screenshot-Email.ps1` | スクリーンショット撮影とメール作成のメイン処理 |
| `Run-Hidden.vbs` | PowerShellスクリプトを非表示で実行するためのVBSラッパー |
| `Setup-Shortcut.ps1` | ショートカットキーを登録（config.jsonから読み込み） |
| `Setup-Shortcut.bat` | ショートカット登録用バッチファイル（ダブルクリック実行可能） |

## セットアップ手順

### 1. 前提条件
- Windows 10/11
- Microsoft Outlook がインストール済み
- PowerShell 5.1 以上

### 2. 初期設定

#### 2.1 設定ファイルの編集
`config.json` を編集して、自分の情報を設定:
```json
{
  "name": "あなたの名前",
  "mailto": "送信先メールアドレス",
  "cc": "CCに追加するメールアドレス",
  "subject": "【業務連絡】 （{name}） {date}",
  "body": "スクリーンショットを添付します。",
  "selectionBorderColor": "Red",
  "shortCutKey": "CTRL+ALT+S"
}
```

**件名テンプレートの変数**:
- `{name}`: config.jsonの`name`に置き換えられる
- `{date}`: 実行日（yyyy/MM/dd形式）に置き換えられる

**枠の色の選択肢**: Red, Blue, Green, Yellow, Orange, Purple, Cyan, Magenta, Lime, Pink, White, Black など

**ショートカットキーの例**: 
- `CTRL+ALT+S` (デフォルト)
- `CTRL+SHIFT+S`
- `CTRL+ALT+P`
- `CTRL+SHIFT+X`

#### 2.2 ショートカットキーの登録

**方法1: バッチファイルを使用（おすすめ）**

`Setup-Shortcut.bat` をダブルクリックするだけでOK。

**方法2: PowerShellで直接実行**

PowerShellを**管理者権限**で開き、以下を実行:
```powershell
cd "$env:USERPROFILE\Outlook_PowerShell"
.\Setup-Shortcut.bat
```

これにより、スタートメニューにショートカットが登録され、config.jsonで設定したショートカットキーで実行可能になる。

**注意**: ショートカットキーを変更した場合は、再度 `Setup-Shortcut.bat` または `Setup-Shortcut.ps1` を実行して登録を更新

## 使用方法

### 基本的な使い方
1. **Ctrl + Alt + S** を押す
2. 画面が半透明の黒いオーバーレイで覆われる
3. **マウスカーソルが十字（+）に変わる**
4. **マウスをドラッグ**して撮影したい範囲を選択（赤枠でリアルタイム表示）
5. マウスボタンを離すと範囲が確定され、スクリーンショット撮影
6. Outlookのメール作成画面が自動的に開く
7. 画像が添付され、件名・宛先が自動入力される
8. 必要に応じて本文を編集し、送信する

**キャンセル方法**: 範囲選択中に **ESCキー** を押す

### 手動実行
ショートカットキーを使わず、直接実行する場合:
```powershell
powershell -ExecutionPolicy Bypass -File ".\Send-Screenshot-Email.ps1"
```

または、非表示で実行:
```cmd
wscript.exe ".\Run-Hidden.vbs"
```

## テクニカルな仕様

### 動作フロー
```
[Ctrl+Alt+S] 
    ↓
[Run-Hidden.vbs] 
    ↓ (非表示でPowerShellを起動)
[Send-Screenshot-Email.ps1]
    ↓
[1. config.json読み込み]
    ↓
[2. 範囲選択UI表示（半透明オーバーレイ）]
    ↓
[3. マウスドラッグで範囲選択]
    ↓
[4. 選択範囲のスクリーンショット撮影]
    ↓
[5. image/[日付]/に保存]
    ↓
[6. Outlookメール作成]
    ↓
[7. 画像を添付してメール画面表示]
```

### 使用している技術
- **PowerShell**: スクリプト実行環境
- **System.Windows.Forms**: UI作成、マウスイベント処理、スクリーン情報取得
- **System.Drawing**: 画像キャプチャ、範囲選択描画、保存
- **Outlook COM Object**: メール自動作成
- **VBScript**: バックグラウンド実行

### 画像保存形式
- **形式**: PNG
- **フォルダ構造**: `image/yyyyMMdd/`（日付別）
- **ファイル名**: `screenshot_HHmmss.png`（時刻別）

## トラブルシューティング

### スクリプトが実行できない
**問題**: 「スクリプトの実行が無効になっています」というエラー

**解決策**: PowerShellの実行ポリシーを変更
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Outlookメールが作成されない
**問題**: Outlookが起動しない、またはCOMエラー

**解決策**:
- Outlookがインストールされているか確認
- Outlookを一度起動して、初期設定を完了させる
- 32bit/64bit版の不一致を確認

### ショートカットキーが動作しない
**問題**: Ctrl+Alt+Sを押しても反応しない

**解決策**:
1. スタートメニューを開く
2. "Screenshot-Email" ショートカットを右クリック
3. プロパティで「ショートカットキー」を確認・再設定
4. 他のアプリケーションとのキーバインド競合を確認

### 画像フォルダが作成されない
**問題**: imageフォルダにスクリーンショットが保存されない

**解決策**:
- スクリプトのあるフォルダへの書き込み権限を確認
- フルパスが正しいか確認

### 範囲選択画面が表示されない
**問題**: ショートカットを押しても範囲選択画面が出ない

**解決策**:
- PowerShellが正しく起動しているか確認
- 他のアプリケーションが全画面表示になっていないか確認
- タスクマネージャーでPowerShellプロセスを確認

### 選択範囲が正しく撮影されない
**問題**: 選択した範囲と違う場所が撮影される

**解決策**:
- Windowsのディスプレイ設定で拡大率が100%になっているか確認
- マルチディスプレイの場合、配置設定を確認

## カスタマイズ

### 範囲選択枠の色変更
`config.json` の `selectionBorderColor` を編集:
```json
{
  "selectionBorderColor": "Blue"
}
```

**利用可能な色**: Red, Blue, Green, Yellow, Orange, Purple, Cyan, Magenta, Lime, Pink, White, Black など（.NET Color構造体の名前付き色）

### メール件名の変更
[Send-Screenshot-Email.ps1](Send-Screenshot-Email.ps1#L31) の以下の行を編集:
```powershell
$mail.Subject = "【業務連絡】 （$($config.name)） $dateFormatted"
```

### 画像サイズの変更
[Send-Screenshot-Email.ps1](Send-Screenshot-Email.ps1#L44) の以下の行で画像幅を調整:
```html
<p><img src="cid:screenshot" width="1200" /></p>
```

### ショートカットキーの変更
`config.json` の `shortCutKey` を編集:
```json
{
  "shortCutKey": "CTRL+SHIFT+S"
}
```

変更後、`Setup-Shortcut.ps1` を再実行してショートカットを更新してください。

**使用可能な修飾キー**: CTRL, SHIFT, ALT
**例**: `CTRL+ALT+P`, `CTRL+SHIFT+X`, `CTRL+ALT+SHIFT+S`
