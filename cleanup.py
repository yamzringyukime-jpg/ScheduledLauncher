#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import win32com.client


def delete_scheduled_task(task_name):
    """タスクスケジューラからタスクを削除"""
    try:
        scheduler = win32com.client.Dispatch('Schedule.Service')
        scheduler.Connect()
        root_folder = scheduler.GetFolder('\\')

        root_folder.DeleteTask(task_name, 0)
        return True
    except Exception as e:
        print(f"エラー: {e}")
        return False


def main():
    task_name = "ScheduledLauncher"

    if len(sys.argv) > 1:
        task_name = sys.argv[1]

    print(f"タスク '{task_name}' を削除中...")

    if delete_scheduled_task(task_name):
        print(f"タスク '{task_name}' を削除しました。")
    else:
        print(f"タスク '{task_name}' の削除に失敗したか、存在しません。")

    input("Enterキーを押して終了してください...")


if __name__ == '__main__':
    main()
