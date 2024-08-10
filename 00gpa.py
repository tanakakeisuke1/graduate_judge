import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
from tkinterdnd2 import DND_FILES, TkinterDnD

def graduate_judge(file_path=None):
    result_text = ""
    try:
        file_path = file_path_entry.get()
        if file_path:
            ''' data1 は元のデータ　dataは合のみのデータ data2はGPA用のデータ'''
            #CSVファイルの読み込み
            data1 = pd.read_csv(file_path, encoding='cp932', header=3)

            #列,行の削除
            data = data1[(data1['合否'] == '合')]

            #GPA計算
            data1['得点'] = pd.to_numeric(data1['得点'], errors='coerce')
            data2 = data1[(data1['合否'].isin(['合', '否'])) & (data1['評語'] != '認定') & (data1['得点'] != '---')].copy()
            data2['GPA'] = data2['単位数'] * ((data2['得点'] - 50) / 10)
            gpa_sum = data2['GPA'].sum()
            tanni_sum2 = data2['単位数'].sum()
            if tanni_sum2 != 0:
                gpa = gpa_sum / tanni_sum2
            else:
                gpa = 0 
            result_text = 'GPA：{:.4f}  '.format(gpa)

            #単位数のカウント
            tanni_sum = data['単位数'].sum()
            result_text += '修得単位数 {}単位'.format(tanni_sum)

            # エラーメッセージ用のラベルを初期化
            error_message_label.config(text="")

        else:
            result_text = "ファイルを選択してください"

    except Exception as e:
        error_message = "ファイルを読み込めませんでした: " + str(e)
        error_message_label.config(text=error_message)

    result_text_widget.delete(1.0, tk.END)
    result_text_widget.insert(tk.END, result_text)

#ファイル名を取得する関数
def get_filename(file_path):
    return file_path.split("/")[-1]

#ファイルをダイアログから選択できるようにする
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        filename = get_filename(file_path)
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, file_path) #絶対パス指定
        file_path_entry2.delete(0, tk.END)
        file_path_entry2.insert(0, filename) #見えるところ

#ファイルがドロップされたとき
def on_drop(event):
    file_path = event.data
    fix = file_path.replace('\\', '/').replace('{', '').replace('}', '') #ファイルパスに日本語が含まれる際に起きるバグへの対処
    filename = get_filename(fix)
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, fix) #絶対パス指定
    file_path_entry2.delete(0, tk.END)
    file_path_entry2.insert(0, filename) #見えるところ

#ファイルがドロップできるとき
def on_enter(event):
    overlay_frame.config(bg='light blue')

#ファイルがドロップできないとき
def on_leave(event):
    overlay_frame.config(bg='white')

root = TkinterDnD.Tk()
root.geometry("1000x530") #初期ウィンドウサイズ
root.title("GPA計算機")

frame = tk.Frame(root)
frame.pack(padx=20, pady=20)

#ドラッグにより色が変わるところ
overlay_frame = tk.Frame(frame, width=500, height=100, bg='white', highlightthickness=2, highlightbackground='grey')
overlay_frame.grid(row=6, column=0, columnspan=2, pady=10)

#1行目
file_path_label = tk.Label(frame, text="１.成績CSVファイルを選択、もしくはドラッグアンドドロップしてください：", font=("Meiryo", 17))
file_path_label.grid(row=0, column=0)

#2行目：ファイル名表示ボックス
file_path_entry = tk.Entry(frame)
file_path_entry.config(font=('Helvetica', 12))
file_path_entry.grid(row=1, column=0, columnspan=2)

file_path_entry2 = tk.Entry(frame)
file_path_entry2.config(font=('Helvetica', 12))
file_path_entry2.grid(row=1, column=0, columnspan=2)

#3行目：ボタン
browse_button = tk.Button(frame, text="ファイルを探す", command=browse_file)
browse_button.grid(row=2, column=0, columnspan=2, pady=1)

#5行目：色が変わるところ
output_label = tk.Label(frame, text="ここにドラッグアンドドロップ")
output_label.grid(row=4, column=0, columnspan=2)
overlay_frame.grid(row=4, column=0, columnspan=2)
overlay_frame.config(bg='white')

#6行目：ボタン
graduate_button = tk.Button(frame, text="\n    ２.進級＆卒業判定する    \n", command=graduate_judge, font=("Meiryo", 10))
graduate_button.grid(row=5, column=0)

#7行目：改行、結果表示
result_text_widget = tk.Text(frame, wrap=tk.WORD, height=8, width=60)
result_text_widget.grid(row=6, column=0)

#7行目：エラーメッセージ
error_message_label = tk.Label(frame, text="", fg="red")
error_message_label.grid(row=7, column=0, columnspan=2)

#注意書き
frame2 = tk.Frame(root)
frame2.pack()
attention_message_label1 = tk.Label(frame2,text="※このアプリはオフライン環境で使用できます。")
attention_message_label2 = tk.Label(frame2,text="成績情報を利用、収集、外部送信、第三者へ提供することはありません。")
attention_message_label3 = tk.Label(frame2,text="あくまで目安なので、このアプリは自己責任でご使用ください。", fg="red")
attention_message_label1.grid(row=1)
attention_message_label2.grid(row=2)
attention_message_label3.grid(row=3)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)
root.dnd_bind('<<DropEnter>>', on_enter)
root.dnd_bind('<<DropLeave>>', on_leave)

root.mainloop()
