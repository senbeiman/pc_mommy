import tkinter as tk
import tkinter.ttk as ttk
from ctypes import *
from datetime import datetime, timedelta
import slackweb
import subprocess

DEBUG_MODE = False  # デバッグモード切替用
SLEEP_WAIT = 10000  # 就寝時刻になってからスリープまでの猶予期間(ms)
CHECK_CYCLE = 10000  # 就寝時刻かどうか確認するサイクル(ms)
AWAKE_TIME = "06:00"  # スリープを解除する時刻
HOSTS_FILEPATH = "C:/Windows/system32/drivers/etc/HOSTS"  # HOSTSファイルのパス
DEFAULT_EXE = "Solitaire.exe"
DEFAULT_HOST = "www.youtube.com"
DEFAULT_WEBHOOK_URL = "https://hooks.slack.com/services/******/"  # Slack通知用URLのデフォルト値


class Application(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.noticed_flag = False  # リマインダ表示済みかどうかの判定
        self.boosting_flag = False  # お仕置きでスリープ解除中か判定
        self.sleeping_flag = False  # 強制スリープ実行時間かどうか判定，お仕置き中でも下りない
        self.sleep_notice_flag = False  # スリープ通知1回しかしないため
        self.penalty_files = []
        self.penalty_sites = []
        self.penalty_endtimes = []
        self.master = master
        self.sleeping_img = tk.PhotoImage(file="sleeping.png").subsample(4)  # メッセージアイコン用画像の読み込み
        self.sleepy_img = tk.PhotoImage(file="sleepy.png").subsample(4)
        self.boost_img = tk.PhotoImage(file="boost.png").subsample(4)
        self.get_id()
        self.create_widgets()
        self.check_admin()
        self.check_time()
    '''
    管理者権限でアプリを起動したかどうかを確認し，違ったら終了する
    '''
    def check_admin(self):
        if not windll.shell32.IsUserAnAdmin():
            subwin_admin = tk.Toplevel(self, padx=5, pady=5)
            subwin_admin.attributes("-topmost", True)
            subwin_admin.title("エラー")
            subwin_admin.update_idletasks()
            subwin_admin.geometry("+{0:}+{1:}".format(str(subwin_admin.winfo_screenwidth() // 2 - subwin_admin.winfo_width() // 2),
                                              str(subwin_admin.winfo_screenheight() // 2 - subwin_admin.winfo_height() // 2)))
            tk.Message(subwin_admin, text="エラー:\n本アプリは管理者権限で実行してください（3秒後に自動的に終了します）", width=160).grid(row=0, column=0)
            self.after(3000, self.main_closing)
    '''
    Slack通知用にPC固有のIDを取得する
    '''
    def get_id(self):
        kernel32 = WinDLL("kernel32")
        kernel32.GetComputerNameW.restype = c_bool
        kernel32.GetComputerNameW.argtypes = (c_wchar_p, POINTER(c_uint32))
        lenComputerName = c_uint32()
        kernel32.GetComputerNameW(None, lenComputerName)
        computerName = create_unicode_buffer(lenComputerName.value)
        kernel32.GetComputerNameW(computerName, lenComputerName)
        self.computer_name = computerName.value
    '''
    メイン画面の作成
    '''
    def create_widgets(self):
        tk.Label(self, text="目標就寝時刻").grid(row=1, column=0, sticky="w")

        self.time_val = tk.StringVar()
        cb = ttk.Combobox(self, textvariable=self.time_val, width=10, state="readonly")

        cb["values"] = ("23:00", "23:30", "00:00", "00:30", "01:00", "01:30", "02:00")
        cb.current(0)
        cb.grid(row=1, column=1, sticky="w")

        tk.Label(self, text="お仕置き設定").grid(row=2, column=0, sticky="w")

        self.penalty_val = tk.StringVar()
        self.penalty_val.set("")
        self.penalty_label = tk.Label(self, textvariable=self.penalty_val)
        self.penalty_label.grid(row=1, column=2, columnspan=2, sticky="w")

        tk.Label(self, text="アプリ").grid(row=3, column=0, sticky="e")
        self.file_var = tk.StringVar()
        self.file_var.set(DEFAULT_EXE)
        tk.Entry(self, textvariable=self.file_var, width=30).grid(row=3, column=1, columnspan=2, sticky="e")
        tk.Button(self, text="テスト", command=self.test_file, width=8).grid(row=3, column=3)

        tk.Label(self, text="Webサイト").grid(row=4, column=0, sticky="e")
        self.site_var = tk.StringVar()
        self.site_var.set(DEFAULT_HOST)
        tk.Entry(self, textvariable=self.site_var, width=30).grid(row=4, column=1, columnspan=2, sticky="e")
        self.sitetest_var = tk.StringVar()
        self.sitetest_var.set("テスト")
        tk.Button(self, textvariable=self.sitetest_var, command=self.test_site, width=8).grid(row=4, column=3, padx=5)

        tk.Label(self, text="Slack通知用 Webhook URL").grid(row=5, column=0,columnspan=2, sticky="w")
        self.webhook_var = tk.StringVar()
        self.webhook_var.set(DEFAULT_WEBHOOK_URL)
        tk.Entry(self, textvariable=self.webhook_var, width=40).grid(row=6, column=0, columnspan=3, sticky="e")
        tk.Button(self, text="テスト", command=self.test_webhook, width=8).grid(row=6, column=3)

        slack = slackweb.Slack(url=self.webhook_var.get())
        slack.notify(text=self.computer_name+"の監視を始めたわよ")

        if DEBUG_MODE:
            tk.Label(self, text="デモ用時刻").grid(row=7, column=0, sticky="w")
            self.clock_var = tk.StringVar()
            self.clock_var.set("2019-6-15 21:00:00")
            tk.Entry(self, textvariable=self.clock_var).grid(row=7, column=1, columnspan=2)
            self.debug_var = tk.IntVar()
            self.debug_var.set(0)
            rd_frame = tk.Frame(self)
            rd_frame.grid(row=8, column=1, columnspan=2)
            tk.Radiobutton(rd_frame,text="1", variable=self.debug_var, value=1, command=self.time_set).grid(row=0, column=0)
            tk.Radiobutton(rd_frame,text="2", variable=self.debug_var, value=2, command=self.time_set).grid(row=0, column=1)
            tk.Radiobutton(rd_frame,text="3", variable=self.debug_var, value=3, command=self.time_set).grid(row=0, column=2)
            tk.Radiobutton(rd_frame,text="4", variable=self.debug_var, value=4, command=self.time_set).grid(row=0, column=3)

        self.grid(column=0, row=0, padx=5, pady=10)
    '''
    デモ用時刻設定
    '''
    def time_set(self):
        val = self.debug_var.get()
        if val == 1:
            self.clock_var.set("2019-6-15 22:00:00")
        elif val == 2:
            self.clock_var.set("2019-6-15 23:00:00")
        elif val == 3:
            self.clock_var.set("2019-6-16 06:00:00")
        elif val == 4:
            self.clock_var.set("2019-6-16 23:00:00")
    '''
    exeファイル強制終了テスト
    '''
    def test_file(self):
        subprocess.call("taskkill /im "+self.file_var.get()+" /F /T", shell=True)
    '''
    Webサイトアクセス制限テスト
    '''
    def test_site(self):
        if self.sitetest_var.get() == "テスト":
            self.sitetest_name = self.site_var.get()
            with open(HOSTS_FILEPATH, mode="a") as f:
                f.write("\n" + "0.0.0.0 " + self.sitetest_name + " # for TEST")
            self.sitetest_var.set("テスト解除")
        elif self.sitetest_var.get() == "テスト解除":
            with open(HOSTS_FILEPATH, mode="r") as f:
                txt = f.read()
            txt = txt.replace("\n" + "0.0.0.0 " + self.sitetest_name + " # for TEST", "")
            with open(HOSTS_FILEPATH, mode="w") as f:
                f.write(txt)
            self.sitetest_var.set("テスト")
    '''
    Slack通知テスト
    '''
    def test_webhook(self):
        slack = slackweb.Slack(url=self.webhook_var.get())
        slack.notify(text="テスト\n")
    '''
    CHECK_CYCLEの時間おきに各処理を行う
    '''
    def check_time(self):
        self.check_sleep_time()
        if len(self.penalty_endtimes) > 0:
            self.check_penalty_time()
            self.kill_running_file()
        self.after(CHECK_CYCLE, self.check_time)
        if DEBUG_MODE:
            print("noticed_flag:",self.noticed_flag,
                  ",boosting_flag:",self.boosting_flag,", sleeping_flag",self.sleeping_flag,
                  ", penalty_files:",self.penalty_files, ", penalty_sites:",self.penalty_sites,
                  ", penalty_endtimes",self.penalty_endtimes)
    '''
    お仕置き終了時刻になったかどうかの監視
    '''
    def check_penalty_time(self):
        if DEBUG_MODE:
            time_now = datetime.strptime(self.clock_var.get(), '%Y-%m-%d %H:%M:%S')
        else:
            time_now = datetime.now()
        if time_now >= self.penalty_endtimes[0]:
            self.enable_access()
    '''
    スリープ実行・解除時刻の確認
    '''
    def check_sleep_time(self):
        if DEBUG_MODE:
            time_now = datetime.strptime(self.clock_var.get(), '%Y-%m-%d %H:%M:%S')
        else:
            time_now = datetime.now()
        time_now = time_now.strftime("%H:%M")
        print(time_now)
        time_sleep = self.time_val.get()
        h, m = [int(i) for i in time_sleep.split(":")]
        h = (h+24-1) % 24
        time_remind = ":".join([str(h).zfill(2), str(m).zfill(2)])  # 目標時刻1時間前を文字列で表現する
        if time_now == time_sleep:
            self.sleeping_flag = True
        if time_now == AWAKE_TIME:
            self.sleeping_flag = False
            self.boosting_flag = False
            self.sleep_notice_flag = False
        if self.sleeping_flag and not self.boosting_flag:
            self.message_sleep()
            self.after(SLEEP_WAIT, self.sleep_PC)
            self.noticed_flag = False
        elif time_now == time_remind and not self.noticed_flag:
            self.message_remind()
    '''
    目標就寝時刻1時間前になったときの子ウィンドウ生成
    '''
    def message_remind(self):
        self.master.deiconify()
        subwin_remind = tk.Toplevel(self, padx=5, pady=5)
        subwin_remind.attributes("-topmost", True)
        subwin_remind.title("リマインダ")
        x = str(self.winfo_rootx())
        y = str(self.winfo_rooty())
        subwin_remind.geometry("+" + x + "+" + y)
        tk.Label(subwin_remind, image=self.sleepy_img).grid(row=0, rowspan=2, column=1)
        tk.Message(subwin_remind, text="あと1時間で寝る時間よ．\nそろそろ寝る準備しなさい", width=160).grid(row=0, column=0)
        tk.Button(subwin_remind, text="はい", command=subwin_remind.destroy).grid(row=1, column=0)
        self.noticed_flag = True
    '''
    目標就寝時刻になったときの子ウィンドウ生成
    '''
    def message_sleep(self):
        self.master.deiconify()
        self.subwin_sleep = tk.Toplevel(self, padx=5, pady=5)
        self.subwin_sleep.attributes("-topmost", True)
        self.subwin_sleep.title("Time up!")
        x = str(self.winfo_rootx())
        y = str(self.winfo_rooty())
        self.subwin_sleep.geometry("+" + x + "+" + y)
        tk.Label(self.subwin_sleep, image=self.sleeping_img).grid(row=0, rowspan=2, column=1)
        tk.Message(self.subwin_sleep, text="時間よ！もう寝なさい．\n10秒後にスリープするからね", width=160).grid(row=0, column=0)
        tk.Button(self.subwin_sleep, text="お仕置きを受けて阻止", command=self.set_penalty).grid(row=1, column=0)
        self.subwin_sleep.after(SLEEP_WAIT, self.subwin_sleep.destroy)
    '''
    お仕置き開始
    '''
    def set_penalty(self):
        self.master.deiconify()
        subwin_penalty = tk.Toplevel(self, padx=5, pady=5)
        subwin_penalty.attributes("-topmost", True)
        subwin_penalty.title("お仕置き")
        x = str(self.winfo_rootx())
        y = str(self.winfo_rooty())
        subwin_penalty.geometry("+" + x + "+" + y)
        self.penalty_files.append(self.file_var.get())
        self.penalty_sites.append(self.site_var.get())
        if DEBUG_MODE:
            time_now = datetime.strptime(self.clock_var.get(), '%Y-%m-%d %H:%M:%S')
        else:
            time_now = datetime.now()
        self.penalty_endtimes.append(time_now+timedelta(days=1))
        tk.Label(subwin_penalty, image=self.boost_img).grid(row=0, rowspan=2, column=1)
        tk.Message(subwin_penalty, text="まだ起きてる気！？\nあんたの大好きな\n「"+self.penalty_files[-1]+"」と\n「"
                                        + self.penalty_sites[-1]+"」\n24時間禁止よ！！",
                   width=160).grid(row=0, column=0)
        tk.Button(subwin_penalty, text="はい", command=subwin_penalty.destroy).grid(row=1, column=0)
        self.boosting_flag = True
        self.disable_access()
        self.penalty_val.set("お仕置き中")
        self.penalty_label.config(fg="white", bg="red")
        self.subwin_sleep.destroy()
        slack = slackweb.Slack(url=self.webhook_var.get())
        slack.notify(text=self.computer_name+"に対して24時間お仕置きを開始したわよ\n"
                          + "- 対象アプリケーション：" + self.penalty_files[-1] + "\n"
                          + "- 対象Webサイト：" + self.penalty_sites[-1] + "\n"
                          + "- 解除時刻：" + self.penalty_endtimes[-1].strftime("%Y-%m-%d %H:%M"))
    '''
    お仕置き対象アプリの強制終了実行
    '''
    def kill_running_file(self):
        for f in self.penalty_files:
            subprocess.call("taskkill /im "+f+" /F /T", shell=True)
    '''
    Webサイトアクセス制限実行
    '''
    def disable_access(self):
        with open(HOSTS_FILEPATH, mode="a") as f:
            f.write("\n" + "0.0.0.0 "+self.penalty_sites[-1]+" #"+self.penalty_endtimes[-1].strftime("%Y-%m-%d %H:%M"))
    '''
    Webサイトアクセス制限解除
    '''
    def enable_access(self):
        with open(HOSTS_FILEPATH, mode="r") as f:
            txt = f.read()
        txt = txt.replace("\n"+"0.0.0.0 "+self.penalty_sites[0]+" #"+self.penalty_endtimes[0].strftime("%Y-%m-%d %H:%M"), "")
        with open(HOSTS_FILEPATH, mode="w") as f:
            f.write(txt)
        slack = slackweb.Slack(url=self.webhook_var.get())
        slack.notify(text=self.computer_name+"へのお仕置きはこれくらいにしといてあげたわよ\n"
                          + "- 対象アプリケーション：" + self.penalty_files[0] + "\n"
                          + "- 対象Webサイト：" + self.penalty_sites[0])
        del self.penalty_sites[0]
        del self.penalty_files[0]
        del self.penalty_endtimes[0]
        if len(self.penalty_endtimes) == 0:
            self.penalty_val.set("")
            self.penalty_label.config(bg='SystemButtonFace')
    '''
    強制スリープ＆ロックの実行
    '''
    def sleep_PC(self):
        if self.sleeping_flag and not self.boosting_flag:
            DISPLAY_OFF = 2
            HWND_BROADCAST = 0xffff
            WM_SYSCOMMAND = 0x0112
            SC_MONITORPOWER = 0xf170

            windll.user32.PostMessageA(HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, DISPLAY_OFF)  # スリープする
            windll.user32.LockWorkStation()  # ロックする
            if not self.sleep_notice_flag:
                slack = slackweb.Slack(url=self.webhook_var.get())
                slack.notify(text=self.computer_name+"をスリープさせておいたわ")  # Slackへ通知
                self.sleep_notice_flag = True
    '''
    メイン画面を閉じるときにSlackへの通知を行うため1クッションはさむ
    '''
    def main_closing(self):
        slack = slackweb.Slack(url=self.webhook_var.get())
        slack.notify(text=self.computer_name+"の監視を終了したわ")
        self.quit()


'''
メイン
'''
def main():
    root = tk.Tk()
    root.title("睡眠管理アプリ　PC母ちゃん")
    app = Application(root)  # アプリケーションインスタンス生成
    root.update_idletasks()
    root.geometry("+{0:}+{1:}".format(str(root.winfo_screenwidth()//2-root.winfo_width()//2),
                                      str(root.winfo_screenheight()//2-root.winfo_height()//2)))  # 画面の中央にウィンドウを表示する
    root.protocol("WM_DELETE_WINDOW", app.main_closing)  # 閉じるボタン押されたらコールバック
    root.mainloop()


if __name__ == "__main__":
    main()
