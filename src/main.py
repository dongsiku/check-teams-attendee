import csv
import re
from pathlib import Path
from getpass import getpass
from tkinter import filedialog
from datetime import datetime

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
        for i, row in enumerate(reader):
            if i == 0:  # headerをパスする
                pass
            temp_attendee_name = format_name(row[0])
            attendees_list.append(temp_attendee_name)
    # print(attendees_list)
    return attendees_list


def read_excel():
    # 初期化する
    EXCEL_FILENAME = Path("研究会用名簿_20211014.xlsx").resolve()
    PASSWD_FILENAME = Path("password.txt")

    # 名簿とパスワードの存在を確認する
    if not EXCEL_FILENAME.exists():
        print("研究会用名簿_20211014.xlsxをダウンロードしてください")
    if PASSWD_FILENAME.exists():
        with open("password.txt", encoding="utf-8") as passwd_f:
            passwd = passwd_f.read().strip()
    else:
        passwd = getpass("Password: ")
        passwd = passwd.strip()

    # Excelファイルと出席者リストを比較し，未確認者の氏名とメールアドレスを取得する
    atendance_list = read_atendance_list(True)
    absentee_list = []
    try:
        excel = win32com.client.Dispatch('Excel.Application')
        workbook = excel.Workbooks.Open(
            EXCEL_FILENAME, False, False, None, passwd
        )
        worksheet = workbook.Worksheets[0]
        mail_list_str = ""
        for i in range(60):
            temp_name = worksheet.Cells.Item(i + 1, 2).Value
            if temp_name is None:
                break
            temp_name = format_name(temp_name)
            do_attend = temp_name in atendance_list
            if do_attend is False:
                # print(temp_name, do_attend)
                temp_mail = worksheet.Cells.Item(i + 1, 4)
                mail_list_str += f"{temp_mail},"
                absentee_list.append(temp_name)
        if len(absentee_list) == 0:
            print("学生全員の出席が確認できました")
        else:
            export_result(absentee_list, mail_list_str)
    finally:
        excel.Quit()


def export_result(absentee_list, mail_list_str):
    absentee_list_str = ""
    for absentee in absentee_list:
        absentee_list_str += f"{absentee.split('+',1)[0]}さん，"
    absentee_list_str = absentee_list_str[0:-1]
    teams_msg = f"現在，{absentee_list_str}の出席が確認できておりません"
    print(f"氏名|\n{absentee_list_str}")
    print(f"メアド|\n{mail_list_str}")
    print(f"Teams message|\n{teams_msg}")
    datetime_str = f"{datetime.now():%Y/%m/%d %H:%M:%S}"
    msg = f"{datetime_str}\n{absentee_list_str}\n{mail_list_str}\n{teams_msg}\n"
    with open("result.txt", "w", encoding="utf-8") as result_f:
        result_f.write(msg)


def format_name(raw_name):
    # 氏名の空白文字の削除とアルファベットの小文字への統一を行う
    formatted_name = raw_name.capitalize()
    formatted_name = re.sub(" |\u3000", "+", formatted_name)
    return formatted_name


if __name__ == "__main__":
    read_excel()
