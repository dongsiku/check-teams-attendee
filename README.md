# 出席者確認アプリ

名簿のExcelファイルと出席者ファイルを比較し，未確認者を抽出する．

## 環境

以下の環境で，動作が確認されている．

- OS: Windows 10, Windows 11
- Python: 3.10.8

動作は確認できていないものの，macOSやLinuxでも使えるはずである．ただし，本文書では基本的にWindows OSでの動作を想定している．

## 初期設定

初期設定を示す．ただし，Pythonがインストールされていない場合は，[このサイト](https://muneue-suwa.github.io/my-site-prototype/docs/python-ja/install-pyenv-win)などを参考に，[公式ホームページ](https://www.python.org/)から，ダウンロード＆インストールしておく．

1. `init.bat`をクリックする．`.env`フォルダが生成されたら成功！ないしは，以下のコマンドを実行する．
    ```powershell
    python -m venv .env
    .env\Scripts\pip install -U pip
    # .env/bin/pip install -U pip  # macOS, Linuxの場合
    .env\Scripts\pip install -r requirements.txt
    # .env/bin/pip install -r requirements.txt  # macOS, Linuxの場合
    ```
2. `check-teams-attendee/`に名簿ファイルを保存する．特にこだわりがなければ，`roster.xlsx`とする（他のファイル名を希望する場合については後述する）．Excelファイルのパスワードは解除しておく．
3. ~~`check-teams-attendee/`に`password.txt`を作成し，パスワードを保存する．~~ （パスワード解除機能は削除されたため，これは不要である）

### 名簿のExcelファイルのフォーマット

表の例を示す．基本的に研究室の名簿と同じであるが，1) ヘッダーを除き区分を分けるための行（Docter，M2など）が削除されていること，2) Teamsか名簿に英字で登録されている学生は，スペルが同じものになっていること（大文字 vs 小文字，全角 vs 半角空白文字の違いはプログラムで対応しているため，変更は不要である）に留意すること．

| 区分  |   氏名    |   フリガナ    |       メールアドレス       |
| :---: | :-------: | :-----------: | :------------------------: |
|  Dr   | 津島 太郎 | ツシマ タロウ | taro.tsushima@example.com  |
|  M2   | 鹿田 花子 | シカダ ハナコ | hanako.shikada@example.com |

なお，以下の情報は無視される．
- ヘッダー（1行目）
- 区分とフリガナ：研究室の名簿ファイルを流用したものであるため，残されているが，プログラムからはこれらの情報を読み込まない．

## 使い方

1. Teamsからダウンロードした主席者ファイルをダウンロードする．これは，会議を立ち上げたアカウントからのみ可能である．
2. `run.bat`をクリックする．あるいは，以下のコマンドを実行する．
    ```powershell
    .env\Scripts\python src\main.py
    # .env/bin/python src/main.py  # macOS, Linuxの場合
    ```
3. ファイルを選択するダイアログが表示されるため，Teamsからダウンロードした主席者ファイルを選択する．

もし，出力に文字化けが発生する場合は，同じ内容が`check-teams-attendee/result.txt`にも保存されるため，これを確認すればよい．

### 任意の名簿ファイルを扱う場合

コマンドライン引数`--roster`を使って指定する．例えば，`my_roster.xlsx`を使用する場合は，以下のように指定する．

```powershell
.env\Scripts\python src\main.py --roster my_roster.xlsx
# .env/bin/python src/main.py --roster my_roster.xlsx  # macOS, Linuxの場合
```

## 再インストール方法

プログラムの保存場所を移動させたときなどに行う．`.env/`を削除した後，初期設定と同じ手順を実行する．

### macOSやLinuxの場合

`.env/`が隠しファイルとして扱われるかもしれない．その場合は，以下のコマンドで削除を行う．

```bash
ls -fl | grep .env  # 存在の有無を調査する
rm -rf .env
ls -fl | grep .env  # 再度，存在の有無を調査する
```

## 本プログラムの問題点

一度入室すれば退出したとしても出席していると判定されてしまうことが，問題として挙げられる．これは，本プログラムでの出席判定が，Teamsから出力される`meetingAttendanceList.csv`での氏名の有無に依存してるためである．入室した記録と退出した記録を区別して認識させることで，退出している学生も判定できるかもしれない．
