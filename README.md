# ScheduledLauncher

Windowsタスクスケジューラと連携した定時起動アプリ。スリープ復帰時にも確実に実行され、タスクトレイ常駐でカレンダーUIによるON/OFF設定が可能。

## 機能

- ブラウザ自動起動（通常/シークレットモード、ブラウザ選択対応：default/chrome/edge/firefox）
- 外部アプリ起動
- 時間差順次起動（複数アプリを指定秒数間隔で順次開く）
- カレンダーUIによる日付ON/OFF設定
- タスクスケジューラ連携（スリープ復帰時トリガー対応）
- タスクトレイ常駐

## 前提条件

- Windows 10/11
- 管理者権限を持つユーザーアカウント（タスクスケジューラ登録に必要）

## 開発環境セットアップ

```bash
pip install -r requirements.txt
```

## ビルド

PyInstallerを使用してEXE化します。

```bash
pyinstaller build.spec
```

ビルド後、以下のフォルダが生成されます：
- `launcher/` - launcher.exe（メインアプリ）
- `cleanup/` - cleanup.exe（タスク手動削除ツール）

配布用にフォルダを整理してください：
- `launcher/launcher.exe` → `launcher.exe`
- `cleanup/cleanup.exe` → `cleanup.exe`
- `config.json` → そのままコピー
- `README.txt` → このファイル

## 使い方

### 通常起動

`launcher.exe` を実行すると、タスクトレイに常駐します。

タスクトレイメニュー：
- 設定を開く - 起動時間、アプリ/URL、カレンダー設定
- 今すぐ実行 - 手動テスト実行
- 終了 - アプリ終了（タスクスケジューラ登録も削除）

### カレンダー設定

設定画面の「カレンダー設定」タブで、休日を設定できます。
- 日付をクリックするとON/OFFが切り替わります
- 灰色の日付は休日（自動起動スキップ）

### タスクスケジューラ手動削除

アプリが異常終了した場合、タスクが残存することがあります。以下の方法で削除してください：

1. **タスクスケジューラGUI**
   - Win+R → `taskschd.msc`
   - タスクスケジューラライブラリ → ScheduledLauncher を右クリック → 削除

2. **コマンドプロンプト**
   ```cmd
   schtasks /delete /tn "ScheduledLauncher" /f
   ```

3. **PowerShell**
   ```powershell
   Unregister-ScheduledTask -TaskName "ScheduledLauncher" -Confirm:$false
   ```

4. **cleanup.exe**
   - 配布フォルダ内の `cleanup.exe` をダブルクリック

## ポータブル化

アプリは完全ポータブル仕様です。フォルダごと持ち運べます。

**注意：**
- タスクスケジューラ登録はシステムに変更を加えます
- 異常終了時はタスクが残存する可能性があります（手動削除が必要）
- 設定ファイル（config.json）はEXEと同じフォルダに保存されます

## 技術スタック

- Python 3.10+
- pystray（タスクトレイ）
- tkinter（カレンダーUI）
- tkcalendar（カレンダーウィジェット）
- pywin32（タスクスケジューラ連携）
- PyInstaller（EXE化）

## ライセンス

MIT License
