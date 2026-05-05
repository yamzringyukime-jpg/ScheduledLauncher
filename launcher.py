#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import json
import subprocess
import webbrowser
from pathlib import Path
import threading
import time
from datetime import datetime, timedelta
import traceback
import winreg
import win32event
import win32api
import winerror

import pystray
from PIL import Image, ImageDraw
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import simpledialog
from tkcalendar import Calendar

import win32com.client


class Logger:
    """ログ出力クラス - 実行ファイルと同階層にlog.txtを出力"""

    def __init__(self):
        # ログファイルのパス（スクリプトまたはEXEと同階層）
        if getattr(sys, 'frozen', False):
            # EXE実行時
            self.log_path = Path(sys.executable).parent / 'log.txt'
        else:
            # スクリプト実行時
            self.log_path = Path(__file__).parent / 'log.txt'

    def write(self, message, level='INFO'):
        """ログを書き込み"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_line = f'[{timestamp}] [{level}] {message}\n'
        try:
            with open(self.log_path, 'a', encoding='utf-8') as f:
                f.write(log_line)
        except Exception as e:
            # ログ書き込み失敗時はコンソールに出力
            print(f'ログ書き込み失敗: {e}')

    def info(self, message):
        """INFOレベルログ"""
        self.write(message, 'INFO')

    def error(self, message):
        """ERRORレベルログ"""
        self.write(message, 'ERROR')

    def warning(self, message):
        """WARNINGレベルログ"""
        self.write(message, 'WARNING')

    def debug(self, message):
        """DEBUGレベルログ"""
        self.write(message, 'DEBUG')


# グローバルLoggerインスタンス
logger = Logger()


class TaskScheduler:
    """タスクスケジューラ操作クラス - Windowsタスクスケジューラと連携"""

    def __init__(self, task_name):
        self.task_name = task_name
        try:
            # タスクスケジューラサービスに接続
            self.scheduler = win32com.client.Dispatch('Schedule.Service')
            self.scheduler.Connect()
            self.root_folder = self.scheduler.GetFolder('\\')
        except Exception as e:
            error_msg = f'タスクスケジューラ接続失敗: {str(e)}\npywin32がインストールされているか確認してください。'
            logger.error(f'{error_msg}\n{traceback.format_exc()}')
            raise RuntimeError(error_msg)

    def task_exists(self):
        """タスクが存在するか確認"""
        try:
            self.root_folder.GetTask(self.task_name)
            return True
        except Exception as e:
            # タスクが存在しない場合は例外が発生する（正常）
            logger.debug(f'タスク存在確認: {self.task_name} - 存在しない')
            return False

    def create_task(self, script_path, trigger_time):
        """タスクを作成 - 毎日指定時刻に実行するタスクを登録"""
        # 既存タスクがある場合は削除してから再登録
        if self.task_exists():
            logger.info(f'既存タスク検出: {self.task_name} - 削除して再登録')
            self.delete_task()

        try:
            task_def = self.scheduler.NewTask(0)

            # トリガー設定（毎日実行）
            trigger = task_def.Triggers.Create(2)  # TASK_TRIGGER_DAILY
            trigger.StartBoundary = trigger_time
            trigger.DaysInterval = 1

            # アクション設定（スクリプト実行）
            action = task_def.Actions.Create(0)  # TASK_ACTION_EXEC
            action.Path = sys.executable
            action.Arguments = f'"{script_path}" --run'
            action.WorkingDirectory = str(Path(script_path).parent)

            # 設定（スリープ復帰時実行）
            task_def.Settings.WakeToRun = True
            task_def.Settings.StartWhenAvailable = True
            task_def.Settings.StopIfGoingOnBatteries = False

            # タスク登録
            self.root_folder.RegisterTaskDefinition(
                self.task_name,
                task_def,
                6,  # TASK_CREATE_OR_UPDATE
                None,  # ユーザー
                None,  # パスワード
                3  # TASK_LOGON_INTERACTIVE_TOKEN (ログオン時のみ実行)
            )
            logger.info(f'タスク作成成功: {self.task_name} - {trigger_time}')
            return True, None
        except Exception as e:
            error_msg = f'タスク作成失敗: {str(e)}\n管理者権限で実行してください。'
            logger.error(f'{error_msg}\n{traceback.format_exc()}')
            return False, error_msg

    def delete_task(self):
        """タスクを削除"""
        try:
            self.root_folder.DeleteTask(self.task_name, 0)
            logger.info(f'タスク削除成功: {self.task_name}')
            return True, None
        except Exception as e:
            error_msg = f'タスク削除失敗: {str(e)}'
            logger.error(f'{error_msg}\n{traceback.format_exc()}')
            return False, error_msg


class ConfigManager:
    """設定ファイル管理クラス - config.jsonの読み書きを担当"""

    def __init__(self, config_path):
        self.config_path = config_path
        self.config = self.load()

    def load(self):
        """設定ファイルを読み込み - ファイル破損時はデフォルト設定で復旧"""
        default_config = {
            "task_name": "ScheduledLauncher",
            "launch_time": "08:30",
            "apps": [],
            "calendar_disabled_dates": [],
            "enabled": True
        }
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                logger.info(f'設定ファイル読み込み成功: {self.config_path}')
                return config
        except FileNotFoundError:
            logger.warning(f'設定ファイル不存在: {self.config_path} - デフォルト使用')
            return default_config
        except json.JSONDecodeError as e:
            logger.error(f'設定ファイル破損: {self.config_path} - デフォルト復旧\n{str(e)}')
            return default_config
        except Exception as e:
            logger.error(f'設定ファイル読み込みエラー: {self.config_path} - {str(e)}\n{traceback.format_exc()}')
            return default_config

    def save(self):
        """設定ファイルを保存"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            logger.info(f'設定ファイル保存成功: {self.config_path}')
        except Exception as e:
            logger.error(f'設定ファイル保存失敗: {self.config_path} - {str(e)}\n{traceback.format_exc()}')


class LauncherApp:
    """メインアプリケーション - タスクトレイ常駐と設定管理"""
    
    @staticmethod
    def get_browser_path():
        """レジストリから各種ブラウザのパスを取得"""
        browser_paths = {}
        paths_to_check = {
            'chrome': r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe',
            'edge': r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe',
            'firefox': r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\firefox.exe'
        }
        
        for name, reg_path in paths_to_check.items():
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_READ)
                path = winreg.QueryValue(key, None)
                winreg.CloseKey(key)
                if path and Path(path).exists():
                    browser_paths[name] = path
            except Exception:
                continue
        return browser_paths

    @staticmethod
    def launch_url(app_data):
        """指定されたブラウザとモードでURLを開く"""
        target = app_data['target']
        browser_choice = app_data.get('browser', 'default')
        incognito = app_data.get('incognito', False)
        
        try:
            if browser_choice == 'default':
                # 既定のブラウザ（シークレット不可）
                if not webbrowser.open(target):
                    raise RuntimeError("既定のブラウザを開けませんでした。")
                return

            paths = LauncherApp.get_browser_path()
            path = paths.get(browser_choice)
            
            if not path:
                # 指定ブラウザが見つからない場合は既定でフォールバック
                logger.warning(f'{browser_choice} が見つからないため既定のブラウザを使用します')
                if not webbrowser.open(target):
                    raise RuntimeError("既定のブラウザを開けませんでした。")
                return

            # 各ブラウザのシークレット用引数
            args = [path]
            if incognito:
                if browser_choice == 'chrome': args.append('--incognito')
                elif browser_choice == 'edge': args.append('--inprivate')
                elif browser_choice == 'firefox': args.append('-private-window')
            
            args.append(target)
            subprocess.Popen(args)
        except Exception as e:
            error_msg = f"URLの起動に失敗しました。\n\n対象: {target}\nブラウザ: {browser_choice}\nエラー: {str(e)}\n\n※指定したブラウザがインストールされているか確認してください。"
            logger.error(error_msg)
            # GUIモードの場合はダイアログで通知
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("起動エラー", error_msg)
            root.destroy()

    def __init__(self):
        # スクリプトパスと設定ファイルパスの取得（PyInstaller対応）
        if getattr(sys, 'frozen', False):
            # EXE実行時
            self.script_path = Path(sys.executable).absolute()
        else:
            # スクリプト実行時
            self.script_path = Path(__file__).absolute()
        self.config_path = self.script_path.parent / 'config.json'
        self.config_manager = ConfigManager(self.config_path)
        self.task_scheduler = TaskScheduler(self.config_manager.config['task_name'])

        logger.info('アプリケーション起動')

        # 起動時クリーンアップ（残存タスクの削除）
        self.cleanup_existing_task()

        # タスクトレイアイコンの作成
        self.icon = None
        self.create_tray_icon()

        # 設定画面の多重起動防止フラグ
        self.settings_window_open = False

    def cleanup_existing_task(self):
        """残存タスクをクリーンアップ - 異常終了時のタスク残存を防止"""
        if self.task_scheduler.task_exists():
            logger.info('残存タスク検出 - 削除実行')
            success, error_msg = self.task_scheduler.delete_task()
            if not success:
                logger.warning(f'クリーンアップ失敗: {error_msg}')
                # ユーザーに通知（起動時なのでダイアログではなくログのみ）

    def create_tray_icon(self):
        """タスクトレイアイコン作成 - 右クリックメニューの定義"""
        # アイコン画像生成
        image = self.create_icon_image()

        # 右クリックメニュー定義
        menu = pystray.Menu(
            pystray.MenuItem('設定を開く', self.open_settings),
            pystray.MenuItem('今すぐ実行', self.run_apps),
            pystray.MenuItem('終了', self.quit_app)
        )

        self.icon = pystray.Icon('ScheduledLauncher', image, 'ScheduledLauncher', menu)
        logger.info('タスクトレイアイコン作成完了')

    def create_icon_image(self):
        """アイコン画像を生成 - 青色背景にSL（ScheduledLauncher）の文字"""
        width = 64
        height = 64
        image = Image.new('RGB', (width, height), color=(0, 120, 215))
        dc = ImageDraw.Draw(image)
        dc.text((10, 20), 'SL', fill=(255, 255, 255))
        return image

    def open_settings(self, icon=None, item=None):
        """設定画面を開く - tkinterによるGUI設定画面"""
        if self.settings_window_open:
            logger.info('設定画面は既に開いています')
            return

        self.settings_window_open = True
        logger.info('設定画面を開く')

        # Tkinterを別スレッドで実行（pystrayのサブスレッドから直接呼ぶと×ボタンが効かないため）
        def run_settings():
            try:
                root = tk.Tk()
                root.title('ScheduledLauncher 設定')
                root.geometry('600x500')

                # 'X'ボタン（閉じる）のフリーズ対策
                def on_closing():
                    try:
                        # mainloop() を終了させ、ウィンドウリソースを破棄
                        root.quit()
                        root.destroy()
                    except Exception as e:
                        # 破棄の際のエラー（Tclエラー等）はログに記録するが、続行
                        logger.debug(f'ウィンドウ破棄中の軽微なエラー: {e}')
                root.protocol("WM_DELETE_WINDOW", on_closing)

                # ノートブック（タブ）
                notebook = ttk.Notebook(root)
                notebook.pack(fill='both', expand=True, padx=10, pady=10)

                # 基本設定タブ
                basic_frame = ttk.Frame(notebook, padding=10)
                notebook.add(basic_frame, text='基本設定')

                # 起動時間
                ttk.Label(basic_frame, text='起動時間:').grid(row=0, column=0, sticky='w', pady=5)
                time_var = tk.StringVar(value=self.config_manager.config['launch_time'])
                ttk.Entry(basic_frame, textvariable=time_var).grid(row=0, column=1, sticky='ew', pady=5)

                # 有効/無効
                enabled_var = tk.BooleanVar(value=self.config_manager.config['enabled'])
                ttk.Checkbutton(basic_frame, text='有効', variable=enabled_var).grid(row=1, column=0, columnspan=2, sticky='w', pady=5)

                # アプリリスト
                ttk.Label(basic_frame, text='起動リスト:').grid(row=2, column=0, columnspan=2, sticky='w', pady=(10, 5))

                apps_frame = ttk.Frame(basic_frame)
                apps_frame.grid(row=3, column=0, columnspan=2, sticky='nsew')

                apps_listbox = tk.Listbox(apps_frame, height=10)
                apps_listbox.pack(side='left', fill='both', expand=True)

                scrollbar = ttk.Scrollbar(apps_frame, orient='vertical', command=apps_listbox.yview)
                scrollbar.pack(side='right', fill='y')
                apps_listbox.config(yscrollcommand=scrollbar.set)

                # アプリリスト更新
                def refresh_apps():
                    apps_listbox.delete(0, tk.END)
                    for app in self.config_manager.config['apps']:
                        app_type = 'URL' if app['type'] == 'url' else 'FILE'
                        incognito_str = ' (密)' if app.get('incognito') else ''
                        browser_str = f" [{app.get('browser', 'default')}]" if app['type'] == 'url' else ""
                        apps_listbox.insert(tk.END, f'[{app_type}]{browser_str} {app["target"]} ({app["delay_seconds"]}s){incognito_str}')

                refresh_apps()

                # ボタン
                btn_frame = ttk.Frame(basic_frame)
                btn_frame.grid(row=4, column=0, columnspan=2, pady=10)

                def add_app():
                    dialog = tk.Toplevel(root)
                    dialog.title('アプリ追加')
                    dialog.geometry('400x350')

                    ttk.Label(dialog, text='タイプ:').pack(pady=5)
                    type_var = tk.StringVar(value='url')
                    ttk.Radiobutton(dialog, text='URL', variable=type_var, value='url').pack()
                    ttk.Radiobutton(dialog, text='EXE/ファイル', variable=type_var, value='exe').pack()

                    ttk.Label(dialog, text='パス（EXE/ファイル）またはURL:').pack(pady=5)
                    target_var = tk.StringVar()
                    ttk.Entry(dialog, textvariable=target_var).pack(fill='x', padx=20)

                    ttk.Label(dialog, text='遅延（秒）:').pack(pady=5)
                    delay_var = tk.IntVar(value=0)
                    ttk.Entry(dialog, textvariable=delay_var).pack()

                    # シークレットモード（URLのみ）
                    incognito_var = tk.BooleanVar(value=False)
                    incognito_check = ttk.Checkbutton(dialog, text='シークレットモード', variable=incognito_var)
                    incognito_check.pack(pady=5)
                    
                    # --- ブラウザ選択を追加 ---
                    ttk.Label(dialog, text='使用ブラウザ:').pack(pady=5)
                    browser_var = tk.StringVar(value='default')
                    browser_combo = ttk.Combobox(dialog, textvariable=browser_var, state='readonly')
                    browser_combo['values'] = ('default', 'chrome', 'edge', 'firefox')
                    browser_combo.pack()

                    # URLタイプ以外はブラウザ選択を無効化
                    def on_type_change(*args):
                        if type_var.get() == 'exe':
                            incognito_check.state(['disabled'])
                            browser_combo.state(['disabled'])
                        else:
                            incognito_check.state(['!disabled'])
                            browser_combo.state(['!disabled'])
                    
                    type_var.trace('w', on_type_change)

                    def save_app():
                        # 入力バリデーション
                        target = target_var.get().strip()
                        if not target:
                            messagebox.showerror('エラー', 'パス/URLを入力してください')
                            return

                        try:
                            delay = int(delay_var.get())
                        except (tk.TclError, ValueError):
                            messagebox.showerror('エラー', '遅延時間には数値を入力してください')
                            return

                        if delay < 0:
                            messagebox.showerror('エラー', '遅延時間は0以上の値を入力してください')
                            return
                        
                        app = {
                            'type': type_var.get(),
                            'target': target,
                            'delay_seconds': delay,
                            'browser': browser_var.get() if type_var.get() == 'url' else 'default',
                            'incognito': incognito_var.get() if type_var.get() == 'url' else False
                        }
                        self.config_manager.config['apps'].append(app)
                        refresh_apps()
                        dialog.destroy()

                    ttk.Button(dialog, text='追加', command=save_app).pack(pady=10)

                def remove_app():
                    selection = apps_listbox.curselection()
                    if selection:
                        del self.config_manager.config['apps'][selection[0]]
                        refresh_apps()

                ttk.Button(btn_frame, text='追加', command=add_app).pack(side='left', padx=5)
                ttk.Button(btn_frame, text='削除', command=remove_app).pack(side='left', padx=5)

                # カレンダー設定タブ
                calendar_frame = ttk.Frame(notebook, padding=10)
                notebook.add(calendar_frame, text='カレンダー設定')

                ttk.Label(calendar_frame, text='休日設定（クリックでON/OFF切替）:').pack(pady=(0, 10))

                # カレンダー
                cal = Calendar(calendar_frame, selectmode='day', 
                               year=datetime.now().year, month=datetime.now().month,
                               background='white', foreground='black', 
                               bordercolor='gray', headersbackground='white', 
                               headersforeground='black', selectbackground='#0078d7',
                               selectothermonth=False,
                               showothermonthdays=False,
                               othermonthbackground='white',
                               othermonthwebackground='white')
                cal.pack(fill='both', expand=True)

                # --- 修正ポイント1: 変数定義の移動 ---
                # 初期描画時に参照できるよう、disabled_dates を関数定義の前に移動します
                disabled_dates = self.config_manager.config['calendar_disabled_dates']

                # --- 修正ポイント2: タグ定義の整理と拡充 ---
                # 当月用
                cal.tag_config('normal_normal_weekday', background='#FFFFFF', foreground='#000000')
                cal.tag_config('normal_normal_saturday', background='#FFFFFF', foreground='#0000FF')
                cal.tag_config('normal_normal_sunday', background='#FFFFFF', foreground='#FF0000')
                cal.tag_config('normal_off_weekday', background='#E0E0E0', foreground='#000000')
                cal.tag_config('normal_off_saturday', background='#E0E0E0', foreground='#0000FF')
                cal.tag_config('normal_off_sunday', background='#E0E0E0', foreground='#FF0000')
                
                # 他月用 (文字色を少し薄くしつつ、視認性は確保)
                cal.tag_config('other_normal_weekday', background='#FFFFFF', foreground='#999999')
                cal.tag_config('other_normal_saturday', background='#FFFFFF', foreground='#6666FF')
                cal.tag_config('other_normal_sunday', background='#FFFFFF', foreground='#FF6666')
                cal.tag_config('other_off_weekday', background='#E0E0E0', foreground='#999999')
                cal.tag_config('other_off_saturday', background='#E0E0E0', foreground='#6666FF')
                cal.tag_config('other_off_sunday', background='#E0E0E0', foreground='#FF6666')
                
                def update_calendar_tags(event=None):
                    """表示されている範囲の曜日にタグを適用"""
                    try:
                        # 全イベントをクリア（既存のタグを消去）
                        for ev_id in cal.get_calevents():
                            cal.calevent_remove(ev_id)
                        
                        # 【修正】get_displayed_month() を使って正しい年月を取得
                        month, year = cal.get_displayed_month()
                        view_date = datetime(year, month, 1).date()
                        
                        # カレンダーの表示範囲をカバーしてタグを適用
                        for d in range(-15, 45): 
                            current = view_date + timedelta(days=d)
                            
                            # 当月以外は処理をスキップ（表示もされないのでタグ付け不要）
                            if current.month != month:
                                continue
                                
                            date_str = current.strftime('%Y-%m-%d')
                            is_current_month = (current.month == month)
                            is_off = (date_str in disabled_dates)
                            weekday = current.weekday()
                            
                            # タグ名の組み立て
                            prefix = 'normal' if is_current_month else 'other'
                            state = 'off' if is_off else 'normal'
                            
                            if weekday == 5: day_type = 'saturday'
                            elif weekday == 6: day_type = 'sunday'
                            else: day_type = 'weekday'
                            
                            tag = f"{prefix}_{state}_{day_type}"
                            
                            # 日付に対してタグを適用
                            cal.calevent_create(current, '', tag)
                    except Exception as e:
                        # エラー内容は log.txt に出力されます
                        logger.error(f'カレンダー描画エラー: {e}')

                update_calendar_tags()
                # 月が変わった時にタグを再適用
                cal.bind("<<CalendarMonthChanged>>", update_calendar_tags)

                # カレンダークリックイベント
                def toggle_date(event):
                    # 【修正】選択された日付を取得
                    date_obj = cal.selection_get()
                    
                    # 【修正】現在表示されている年月を取得
                    month, year = cal.get_displayed_month()
                    
                    # 当月以外の日付がクリックされた場合は無視する
                    if date_obj.month != month or date_obj.year != year:
                        logger.debug(f'当月以外の日付クリックを無視: {date_obj}')
                        return

                    date_str = date_obj.strftime('%Y-%m-%d')

                    if date_str in disabled_dates:
                        disabled_dates.remove(date_str)
                    else:
                        disabled_dates.append(date_str)
                    update_calendar_tags()

                cal.bind('<<CalendarSelected>>', toggle_date)

                # 凡例ラベルの追加
                legend_frame = ttk.Frame(calendar_frame)
                legend_frame.pack(fill='x', pady=5)
                ttk.Label(legend_frame, text='凡例:').pack(side='left', padx=5)
                tk.Label(legend_frame, text=' 平日 ', fg='black', bg='white').pack(side='left')
                tk.Label(legend_frame, text=' 土曜 ', fg='blue', bg='white').pack(side='left')
                tk.Label(legend_frame, text=' 日曜 ', fg='red', bg='white').pack(side='left')
                tk.Label(legend_frame, text=' 休日(OFF) ', fg='black', bg='#e0e0e0').pack(side='left')

                # 保存ボタン
                def save_settings():
                    # 起動時間のバリデーション
                    time_str = time_var.get().strip()
                    try:
                        datetime.strptime(time_str, '%H:%M')
                    except ValueError:
                        messagebox.showerror('エラー', '起動時間はHH:MM形式で入力してください（例: 08:30）')
                        return
                    
                    self.config_manager.config['launch_time'] = time_str
                    self.config_manager.config['enabled'] = enabled_var.get()
                    self.config_manager.config['calendar_disabled_dates'] = disabled_dates
                    
                    if not self.config_manager.config['apps']:
                        messagebox.showwarning('確認', '起動リストが空ですが保存しますか？\n(時刻になっても何も起動しません)')

                    try:
                        self.config_manager.save()
                        # タスクスケジューラを再登録（設定反映）
                        if self.config_manager.config['enabled']:
                            self.register_task(parent_root=root)
                        else:
                            # 無効化時はタスクを削除
                            success, error_msg = self.task_scheduler.delete_task()
                            if not success:
                                messagebox.showwarning('警告', f'タスク削除失敗: {error_msg}')
                        messagebox.showinfo('完了', '設定を保存しました')
                        on_closing()
                    except Exception as e:
                        messagebox.showerror('エラー', f'設定保存失敗: {str(e)}')
                        logger.error(f'設定保存失敗: {str(e)}\n{traceback.format_exc()}')

                ttk.Button(root, text='保存', command=save_settings).pack(pady=10)

                root.mainloop()
            except Exception as e:
                logger.error(f'設定画面実行エラー: {e}')
            finally:
                # どのような理由でスレッドが終了しても、必ずフラグをリセットする
                self.settings_window_open = False
                logger.info('設定画面スレッド終了')

        # スレッドで実行
        thread = threading.Thread(target=run_settings, daemon=True)
        thread.start()

    def run_apps(self, icon=None, item=None):
        """アプリを実行 - 設定されたアプリ/URLを順次起動"""
        if not self.config_manager.config['enabled']:
            logger.info('アプリ実行スキップ: 無効設定')
            return

        # アプリリストが空の場合はスキップ
        if not self.config_manager.config['apps']:
            logger.info('アプリ実行スキップ: アプリリストが空')
            msg_root = tk.Tk()
            msg_root.withdraw()
            messagebox.showinfo('情報', '実行するアプリが設定されていません')
            msg_root.destroy()
            return

        # 休日チェック - 今日が休日リストに含まれる場合は実行しない
        today = datetime.now().strftime('%Y-%m-%d')
        if today in self.config_manager.config['calendar_disabled_dates']:
            logger.info(f'アプリ実行スキップ: 休日設定 - {today}')
            return

        logger.info(f'アプリ実行開始: {len(self.config_manager.config["apps"])}個のアプリ')

        # 累積待機時間の計算（順次起動）
        total_delay = 0
        for app in self.config_manager.config['apps']:
            # クロージャ問題回避のためdefault引数を使用
            def execute(app_data):
                try:
                    if app_data['type'] == 'url':
                        LauncherApp.launch_url(app_data)
                        logger.info(f'URL起動: {app_data["target"]} ({app_data.get("browser")})')
                    elif app_data['type'] == 'exe':
                        # EXEや関連付けられたファイルを開く
                        try:
                            os.startfile(app_data['target'])
                            logger.info(f'ファイル起動: {app_data["target"]}')
                        except Exception as e:
                            error_msg = f'起動に失敗しました。\n\n対象パス: {app_data["target"]}\n\nエラー詳細: {str(e)}\n\n※パスが正しいか、または管理者権限が必要でないか確認してください。'
                            logger.error(error_msg)
                            root = tk.Tk()
                            root.withdraw()
                            messagebox.showerror('起動エラー', error_msg)
                            root.destroy()
                except Exception as e:
                    logger.error(f'起動失敗: {app_data["target"]} - {str(e)}')

            # 累積待機で順次起動
            if total_delay > 0:
                logger.info(f'累積待機実行: {app["target"]} - {total_delay}秒後')
                timer = threading.Timer(total_delay, execute, args=[app])
                timer.start()
            else:
                execute(app)
            
            # 次のアプリのために遅延時間を累積
            total_delay += app['delay_seconds']
        
        # GUIモード（タスクトレイ）ではjoinしない（フリーズ回避）
        # --runモード（main関数）でのみjoinする

    def register_task(self, parent_root=None):
        """タスクスケジューラに登録 - 毎日指定時刻に実行するタスクを登録"""
        if not self.config_manager.config['enabled']:
            logger.info('タスク登録スキップ: 無効設定')
            return

        try:
            # 時刻フォーマット変換 (HH:MM -> YYYY-MM-DDTHH:MM:SS)
            time_str = self.config_manager.config['launch_time']
            today = datetime.now().date()
            trigger_datetime = datetime.combine(today, datetime.strptime(time_str, '%H:%M').time())
            
            # 過去の時刻の場合は翌日に設定（即時暴走起動回避）
            if trigger_datetime < datetime.now():
                trigger_datetime += timedelta(days=1)
                logger.info(f'過去の時刻設定検出: 翌日の日付を使用 - {trigger_datetime}')
            
            trigger_time = trigger_datetime.isoformat()

            success, error_msg = self.task_scheduler.create_task(str(self.script_path), trigger_time)
            if not success:
                # エラーダイアログ表示（多重Root回避）
                if parent_root:
                    messagebox.showerror('タスク登録失敗', error_msg)
                else:
                    root = tk.Tk()
                    root.withdraw()
                    messagebox.showerror('タスク登録失敗', error_msg)
                    root.destroy()
        except Exception as e:
            error_msg = f'タスク登録エラー: {str(e)}'
            logger.error(f'{error_msg}\n{traceback.format_exc()}')
            # エラーダイアログ表示（多重Root回避）
            if parent_root:
                messagebox.showerror('タスク登録失敗', error_msg)
            else:
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror('タスク登録失敗', error_msg)
                root.destroy()

    def quit_app(self, icon=None, item=None):
        """アプリ終了 - タスク削除とアイコン停止"""
        logger.info('アプリ終了処理開始')
        # タスク削除
        success, error_msg = self.task_scheduler.delete_task()
        if not success:
            # 手動削除ツールのパスを表示（PyInstaller対応）
            if getattr(sys, 'frozen', False):
                cleanup_path = Path(sys.executable).parent / 'cleanup.exe'
            else:
                cleanup_path = Path(__file__).parent / 'cleanup.py'
            
            root = tk.Tk()
            root.withdraw()
            messagebox.showwarning(
                'タスク削除失敗',
                f'{error_msg}\n\n手動で削除してください：\n{cleanup_path}\n\nまたはコマンドプロンプトで：\nschtasks /delete /tn "ScheduledLauncher" /f'
            )
            root.destroy()
        # アイコン停止
        if self.icon:
            self.icon.stop()
        sys.exit(0)

    def run(self):
        """メインループ - タスク登録後、タスクトレイアイコンを常駐"""
        # タスク登録
        self.register_task()
        # アイコン実行（常駐開始）
        logger.info('タスクトレイ常駐開始')
        self.icon.run()


def main():
    """メイン関数 - 二重起動防止とモード切り替え"""

    # --- コマンドライン引数チェック（Mutexより先に判定する） ---
    # スクリプト実行時: sys.argv[1] == '--run'
    # EXE実行時: sys.argv[2] == '--run'（PyInstallerの仕様）
    # ※ --run モード（タスクスケジューラからの呼び出し）は
    #    常駐プロセスとは別の短命プロセスなので、Mutex による
    #    二重起動チェックの対象外とする
    is_run_mode = False
    if getattr(sys, 'frozen', False):
        # EXE実行時
        if len(sys.argv) > 2 and sys.argv[2] == '--run':
            is_run_mode = True
    else:
        # スクリプト実行時
        if len(sys.argv) > 1 and sys.argv[1] == '--run':
            is_run_mode = True

    # --- 二重起動チェック（常駐モードのみ適用） ---
    # --run モードはタスクスケジューラから呼ばれる短命プロセスなので
    # Mutex チェックをスキップする（常駐プロセスと共存する必要がある）
    mutex = None
    if not is_run_mode:
        mutex_name = "Global\\ScheduledLauncher_Unique_Mutex_Name"
        try:
            # ミューテックスを作成
            mutex = win32event.CreateMutex(None, False, mutex_name)
            # 既に存在するかチェック
            if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
                # 既に起動している場合は終了
                logger.info('二重起動検出 - 既存プロセスが実行中のため終了')
                sys.exit(0)
        except Exception as e:
            # ミューテックス作成失敗時は念のため続行
            logger.warning(f'Mutex作成失敗（続行します）: {e}')
            mutex = None
    
    if is_run_mode:
        # タスクスケジューラからの呼び出し（アプリ実行モード）
        logger.info('タスクスケジューラからの呼び出し検出 - アプリ実行モード')
        # config.jsonのパス（PyInstaller対応）
        if getattr(sys, 'frozen', False):
            config_path = Path(sys.executable).parent / 'config.json'
        else:
            config_path = Path(__file__).parent / 'config.json'
        config_manager = ConfigManager(config_path)
        
        # アプリ実行（休日チェック含む）
        if not config_manager.config['enabled']:
            logger.info('アプリ実行スキップ: 無効設定')
            sys.exit(0)
        
        # アプリリストが空の場合はスキップ
        if not config_manager.config['apps']:
            logger.info('アプリ実行スキップ: アプリリストが空')
            sys.exit(0)
        
        # 休日チェック
        today = datetime.now().strftime('%Y-%m-%d')
        if today in config_manager.config['calendar_disabled_dates']:
            logger.info(f'アプリ実行スキップ: 休日設定 - {today}')
            sys.exit(0)
        
        # アプリ実行
        # 累積待機時間の計算（順次起動）
        total_delay = 0
        for app in config_manager.config['apps']:
            def execute(app_data):
                try:
                    if app_data['type'] == 'url':
                        # クラスメソッドとして追加した launch_url を呼び出す
                        LauncherApp.launch_url(app_data)
                        logger.info(f'URL起動成功: {app_data["target"]} ({app_data.get("browser", "default")})')
                    
                    elif app_data['type'] == 'exe':
                        # EXEの場合はこれまで通り os.startfile
                        try:
                            os.startfile(app_data['target'])
                            logger.info(f'ファイル起動成功: {app_data["target"]}')
                        except Exception as e:
                            logger.error(f'ファイル起動失敗: {app_data["target"]} - {str(e)}')
                        
                except Exception as e:
                    logger.error(f'アプリ起動失敗: {app_data["target"]} - {str(e)}')
            
            # 累積待機で順次起動
            if total_delay > 0:
                logger.info(f'累積待機実行: {app["target"]} - {total_delay}秒後')
                timer = threading.Timer(total_delay, execute, args=[app])
                timer.start()
            else:
                execute(app)
            
            # 次のアプリのために遅延時間を累積
            total_delay += app['delay_seconds']

        # 全てのタイマーがセットされたら、最後のタイマーが発火するまで待機
        if total_delay > 0:
            logger.info(f'全タスクの予約完了。完了まで約{total_delay}秒待機します。')
            time.sleep(total_delay + 1) # 余裕を持って待機

        sys.exit(0)

    # 通常起動（タスクトレイ常駐モード）
    launcher = LauncherApp()
    launcher.run()

    # mutexを保持（GCで削除されないように）
    if mutex is not None:
        pass


if __name__ == '__main__':
    main()
