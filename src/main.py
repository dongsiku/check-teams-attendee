import csv
import re
from pathlib import Path
from tkinter import filedialog
from datetime import datetime
import argparse

import openpyxl


class CheckTeamsAttendee:
    def __init__(self):
        """初期化する
        """
        self.PROJ_DIRNAME = Path(__file__).resolve().parents[1]
        self.DEFAULT_ROSTER_FILENAME = self.PROJ_DIRNAME / "roster.xlsx"
        self.RESULT_FILENAME = self.PROJ_DIRNAME / "result.txt"

    def main(self):
        """Main script

        Raises:
            FileNotFoundError: 名簿のエクセルファイルが見つからなかったとき
            FileNotFoundError: ファイルの選択がキャンセルされたとき
        """
        # Debug modeを判定する
        parser = argparse.ArgumentParser(
            description="Run program with debug mode."
        )
        parser.add_argument(
            "--debug", help="Debug mode", action="store_true",
        )
        parser.add_argument(
            "--roster", type=Path, help="Roster filename",
        )
        args = parser.parse_args()
        IS_DEBUG_MODE = args.debug
        ROSTER_FILENAME = (
            args.roster if args.roster else self.DEFAULT_ROSTER_FILENAME
        )

        # 名簿ファイルの存在を確認する
        if not ROSTER_FILENAME.exists():
            raise FileNotFoundError(
                f"Could not be found: {ROSTER_FILENAME}"
            )

        # 出席者リストのファイルを選択する
        # Debug modeのときは，./meetingAttendanceList.csvを用いる
        meeting_attendance_list_csv = (
            self.PROJ_DIRNAME / "meetingAttendanceList.csv"
        )
        if IS_DEBUG_MODE is False:
            idir = "~/Downloads"
            filetype = [("出席者リスト", "*.csv")]
            temp_meeting_attendance_list_csv = filedialog.askopenfilename(
                filetypes=filetype, initialdir=idir
            )
            if temp_meeting_attendance_list_csv == "":
                raise FileNotFoundError("Canceled")
            meeting_attendance_list_csv = Path(
                temp_meeting_attendance_list_csv
            )

        # 出席者リストを取得する
        attendees_list = self.get_attendees_list_from_csv(
            meeting_attendance_list_csv
        )
        self.collate_attendees_with_roster(attendees_list, ROSTER_FILENAME)

    def get_attendees_list_from_csv(self, meeting_attendance_list_csv: Path):
        """出席者リストをCSVファイルから取得する

        Args:
            meeting_attendance_list_csv (Path): meetingAttendanceList.csvのパス

        Returns:
            list[str]: 出席者リスト
        """
        # 出席者リストのファイルからファイルを作成する
        attendees_list = []
        with meeting_attendance_list_csv.open(
            encoding="utf-16"
        ) as meeting_attendance_list_f:
            reader = csv.reader(meeting_attendance_list_f, delimiter="\t")
            for i, row in enumerate(reader):
                if i == 0:  # headerをパスする
                    continue
                elif len(row) == 0:
                    continue
                temp_attendee_name = self.format_name(row[0])
                attendees_list.append(temp_attendee_name)
        return attendees_list

    def collate_attendees_with_roster(
        self, attendees_list: list[str], roster_filename: Path,
    ):
        """名簿のエクセルファイルを読む

        Args:
            attendees_list (list[str]): 出席者リスト
        """
        # Excelファイルと出席者リストを比較し，未確認者の氏名とメールアドレスを取得する
        absentees_list = []
        workbook = openpyxl.load_workbook(roster_filename)
        worksheet = workbook.worksheets[0]
        email_list_str = ""
        for row in worksheet.values:
            _, temp_name, _, email_address = row[0:4]
            if temp_name is None:
                break
            formatted_name = self.format_name(temp_name)
            do_attend = formatted_name in attendees_list
            if do_attend is False:
                # print(temp_name, do_attend)
                email_list_str += f"{email_address},"
                absentees_list.append(formatted_name)
        if len(absentees_list) == 0:
            print("学生全員の出席が確認できました")
        else:
            self.export_result(absentees_list, email_list_str)

    def export_result(
        self, absentees_list: list[str], email_list_str: list[str]
    ):
        """欠席者リストを出力する

        Args:
            absentees_list (list[str]): 欠席者リスト
            email_list_str (list[str]): 欠席者のメールアドレス
        """
        absentees_list_str = ""
        for absentee in absentees_list:
            absentees_list_str += f"{absentee.split('+',1)[0]}さん，"
        absentees_list_str = absentees_list_str[0:-1]
        teams_msg = f"現在，{absentees_list_str}の出席が確認できておりません"
        print(f"氏名|\n{absentees_list_str}")
        print(f"メアド|\n{email_list_str}")
        print(f"Teams message|\n{teams_msg}")
        datetime_str = f"{datetime.now():%Y/%m/%d %H:%M:%S}"
        msg = f"{datetime_str}\n{absentees_list_str}\n"
        msg += f"{email_list_str}\n{teams_msg}\n"
        with self.RESULT_FILENAME.open("w", encoding="utf-8") as result_f:
            result_f.write(msg)

    def format_name(self, raw_name: str):
        """氏名の空白文字の削除とアルファベットの大小文字の統一を行う．
        これは，Teams上で英字で氏名が登録されている人に対応するためのものである．

        Args:
            raw_name (str): もとの名前の文字列

        Returns:
            str: 整形された名前の文字列
        """
        # 先頭の1文字目を大文字，他を小文字に変換する
        formatted_name = raw_name.capitalize()
        # 空白文字を+に変換する
        formatted_name = re.sub(" |\u3000", "+", formatted_name)

        return formatted_name


if __name__ == "__main__":
    check_teams_attendee = CheckTeamsAttendee()
    check_teams_attendee.main()
