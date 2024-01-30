#!/usr/bin/env python
# coding: utf-8

import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

mail.to = 'yawatanabe@micron.com;'
mail.cc = 'mosano@micron.com;ofushitani@micron.com;'
mail.subject = 'Daily_Report'
mail.bodyFormat = 1
mail.body = '''本日の作業日報を送付します。

1. タスク内容を記述する

明日の予定

'''
# mail app が起動する（内容確認）
mail.display(True)
