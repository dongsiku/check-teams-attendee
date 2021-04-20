import csv
import re
from pathlib import Path
from getpass import getpass
from tkinter import filedialog

import win32com.client


def read_atendance_list(do_forcibly_open_gui_dialog=False):
    # 出席者リストのファイルを選択する
    meeting_attendance_list_csv = Path("meetingAttendanceList.csv")
    if do_forcibly_open_gui_dialog is True or not meeting_attendance_list_csv.exists():
        idir = "~/Downloads"
        filetype = [("出席者リスト", "*.csv")]
        meeting_attendance_list_csv = filedialog.askopenfilename(
            filetypes=filetype, initialdir=idir
        )
        if meeting_attendance_list_csv == "":
            raise FileNotFoundError("Canceled")

    # 出席者リストのファイルからファイルを作成する
    attendees_list = []
    with open(meeting_attendance_list_csv, encoding="utf-16") as meeting_attendance_list_f:
        reader = csv.reader(meeting_attendance_list_f, delimiter="\t")
        for row in reader:
            temp_attendee_name = format_name(row[0])
            attendees_list.append(temp_attendee_name)
    # print(attendees_list)
    return attendees_list


def read_excel():
    # 初期化する
    atendance_list = read_atendance_list(True)
    EXCEL_FILENAME = Path("研究会用名簿_20210420.xlsx").resolve()
    PASSWD_FILENAME = Path("password.txt")

    # 名簿とパスワードの存在を確認する
    if not EXCEL_FILENAME.exists():
        print("研究会用名簿_20210420.xlsxをダウンロードしてください")
    if PASSWD_FILENAME.exists():
        with open("password.txt", encoding="utf-8") as passwd_f:
            passwd = passwd_f.read().strip()
    else:
        passwd = getpass("Password: ")

    # Excelファイルと出席者リストを比較し，未確認者の氏名とメールアドレスを取得する
    try:
        excel = win32com.client.Dispatch('Excel.Application')
        workbook = excel.Workbooks.Open(
            EXCEL_FILENAME, False, False, None, passwd
        )
        worksheet = workbook.Worksheets[0]
        mail_list = ""
        name_list = ""
        for i in range(60):
            temp_name = worksheet.Cells.Item(i + 1, 2).Value
            if temp_name is None:
                break
            temp_name = format_name(temp_name)
            do_attend = temp_name in atendance_list
            if do_attend is False:
                # print(temp_name, do_attend)
                temp_mail = worksheet.Cells.Item(i + 1, 4)
                mail_list += f"{temp_mail},"
                name_list += f"{temp_name},"
        print(name_list)
        print(mail_list)
    finally:
        excel.Quit()


def format_name(name):
    # 氏名の空白文字の削除とアルファベットの小文字への統一を行う
    name = name.lower()
    name = re.sub(" |\u3000", "", name)
    return name


if __name__ == "__main__":
    read_excel()
