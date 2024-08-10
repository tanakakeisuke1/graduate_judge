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
            result_text += '修得単位数 {}単位/124単位中'.format(tanni_sum)

            #教養系科目
            a = data[data['科目区分'] == '大学学習法（大学学習法）']
            a_sum = a['単位数'].sum()

            b = data[(data['科目区分'] == '英語（AR，AL，AW，基礎英語）')]
            b_sum = b['単位数'].sum()
                    
            b1 = data[(data['科目区分'] == '英語（英語）') | (data['科目区分'] == '英語（実践英語）')]
            b1_sum = b1['単位数'].sum()

            c = data[data['科目区分'] == '初修外国語']
            c_sum = c['単位数'].sum()

            d = data[data['科目区分'] == '健康・スポーツ（体育実技）']
            d_sum = d['単位数'].sum()

            e = data[data['科目区分'] == '健康・スポーツ']
            e_sum = e['単位数'].sum()

            f = data[(data['科目区分'] == '情報リテラシー／自然系共通専門基礎') & (data['必選区分'] == '必修')] 
            f_sum = f['単位数'].sum()

            g = data[(data['科目区分'] == '情報リテラシー／自然系共通専門基礎') & (data['必選区分'] == '選択')] 
            g_sum = g['単位数'].sum()

            h = data[(data['科目区分'] == '自然科学（理学）') | (data['科目区分'] == '自然科学（工学）') | (data['科目区分'] == '自然科学（農学）')]
            h_sum = h['単位数'].sum()

            i = data[data['科目区分'] == '人文社会・教育科学']
            i_sum = i['単位数'].sum()

            j = data[(data['科目区分'] == '新潟大学個性化科目（地域入門）') | (data['科目区分'] == '新潟大学個性化科目（地域研究）') | (data['科目区分'] == '新潟大学個性化科目（自由主題）')]
            j_sum = j['単位数'].sum()

            k = data[(data['科目区分'] == '医歯学（医学）') | (data['科目区分'] == '医歯学（歯学）')]
            k_sum = k['単位数'].sum()

            l = data[(data['科目区分'] == '留学生基本科目(日本語)') | (data['科目区分'] == '留学生基本科目(日本事情)')]
            l_sum = l['単位数'].sum()

            #学部専門系科目
            m = data[data['科目区分'] == '専門基礎科目群（選択必修）']
            m_sum = m['単位数'].sum()

            n = data[(data['科目区分'] == '専門応用科目群（必修）') | (data['科目区分'] == '専門応用科目群（必修・工学科共通）')]
            n_sum = n['単位数'].sum()

            o = data[data['科目区分'] == '専門応用科目群（選択必修）']
            o_sum = o['単位数'].sum()

            p = data[data['科目区分'] == '専門応用科目群（選択）']
            p_sum = p['単位数'].sum()

            q = data[data['科目区分'] == '専門応用科目群（特殊選択）']
            q_sum = q['単位数'].sum()

            r = data[data['科目区分'] == '専門応用科目群（自由）']
            r_sum = r['単位数'].sum()

            #CSV破損検知
            all1_sum = a_sum + b_sum + b1_sum + c_sum + d_sum + e_sum + f_sum + g_sum + h_sum + i_sum + j_sum + k_sum + l_sum + m_sum + n_sum + o_sum + p_sum + q_sum + r_sum
            
            s = data[data['科目区分'].notnull()]
            s_sum = s['単位数'].sum()
            specialtanni = s_sum - all1_sum

            #判定 [細分区,必要な単位数,取得した単位数]
            n_count = 0 #卒業
            n_n_count = 0 #進級

            #教養必修
            if a_sum >= 2:
                a_lst = ['教養系','大学学習法',2,2,2]
                a_sum = a_sum - 2
                n_count = n_count + 1
                n_n_count = n_n_count + 1
            else:
                a_lst = ['教養系','大学学習法',2,2,a_sum]
                a_sum = 0

            if b_sum >= 2:
                b_lst = ["",'英語',2,2,2]
                b_sum = b_sum - 2
                n_count = n_count + 1
                n_n_count = n_n_count + 1
            else:
                b_lst = ["",'英語',2,2,b_sum]
                b_sum = 0

            if c_sum >= 2:
                c_lst = ["",'初修外国語',2,2,2]
                c_sum = c_sum - 2
                n_count = n_count + 1
                n_n_count = n_n_count + 1
            else:
                c_lst = ["",'初修外国語',2,2,c_sum]
                c_sum = 0

            if d_sum >= 1:
                d_lst = ["",'健康・スポーツ(体育実技)',1,1,1]
                d_sum = d_sum - 1
                n_count = n_count + 1
                n_n_count = n_n_count + 1
            else:
                d_lst = ["",'健康・スポーツ(体育実技)',1,1,d_sum]
                d_sum = 0

            if f_sum >= 2:
                f_lst = ["",'情報リテラシー',2,2,2]
                f_sum = f_sum - 2
                n_count = n_count + 1
                n_n_count = n_n_count + 1
            else:
                f_lst = ["",'情報リテラシー',2,2,f_sum]
                f_sum = 0

            if i_sum >= 4:
                i_lst = ["",'人文社会・教育科学',4,4,4]
                i_sum = i_sum - 4
                n_count = n_count + 1
                n_n_count = n_n_count + 1
            else:
                i_lst = ["",'人文社会・教育科学',4,4,i_sum]
                i_sum = 0

            #教養選択必修
            fgh_sum = f_sum + g_sum + h_sum
            if fgh_sum  >= 10:
                fgh_lst = ["",'情報リテラシー／自然系共通専門基礎／自然科学',10,10,10]
                fgh_sum = fgh_sum -10
                n_count = n_count + 1
                n_n_count = n_n_count + 1
            else:
                fgh_lst = ["",'情報リテラシー／自然系共通専門基礎／自然科学',10,10,fgh_sum]
                fgh_sum = 0

            ij_sum = i_sum + j_sum
            if ij_sum >= 2:
                ij_lst = ["",'人文社会・教育科学／新潟大学個性化科目',2,2,2]
                ij_sum = ij_sum - 2
                n_count = n_count + 1
                n_n_count = n_n_count + 1
            else:
                ij_lst = ["",'人文社会・教育科学／新潟大学個性化科目',2,2,ij_sum]
                ij_sum = 0

            #教養選択
            sentaku_sum = a_sum + b_sum + b1_sum + c_sum + d_sum + e_sum + fgh_sum + ij_sum +k_sum + l_sum
            if sentaku_sum >= 11:
                sentaku_lst = ["",'選択',11,7,11]
                n_count = n_count + 1
                n_n_count = n_n_count + 1
                sentaku_sum = sentaku_sum -11
            elif (sentaku_sum >= 7) & (sentaku_sum < 11):
                sentaku_lst = ["",'選択',11,7,sentaku_sum]
                sentaku_sum = 0
                n_n_count = n_n_count + 1
            else:
                sentaku_lst = ["",'選択',11,7,sentaku_sum]
                sentaku_sum = 0

            #学部専門系科目
            m_count = 0 #卒業
            m_m_count = 0   #進級

            if m_sum >= 10:
                m_lst = ["学部専門系",'専門基礎科目群B科目',10,10,10]
                m_sum = m_sum - 10
                m_count = m_count + 1
                m_m_count = m_m_count + 1
            else:
                m_lst =  ["学部専門系",'専門基礎科目群B科目',10,10,m_sum]
                m_sum = 0

            if n_sum >= 27:
                n_lst = ["",'専門応用科目群A科目',27,19,27]
                n_sum = n_sum - 27
                m_count = m_count + 1
                m_m_count = m_m_count + 1
            elif (n_sum >= 19) & (n_sum < 27):
                n_lst = ["",'専門応用科目群A科目',27,19,n_sum]
                n_sum = 0
                m_m_count = m_m_count + 1
            else:
                n_lst =  ["",'専門応用科目群A科目',27,19,n_sum]
                n_sum = 0

            mo_sum = m_sum + o_sum

            if mo_sum >= 28:
                mo_lst = ["",'専門応用科目群B科目',28,26,28]
                mo_sum = mo_sum - 28
                m_count = m_count + 1
                m_m_count = m_m_count + 1
            elif (mo_sum >= 26) & (mo_sum < 28):
                mo_lst = ["",'専門応用科目郡B科目',28,26,mo_sum]
                mo_sum = 0
                m_m_count = m_m_count + 1
            else:
                mo_lst =  ["",'専門応用科目群B科目',28,26,mo_sum]
                mo_sum = 0
    
            sentaku2_sum = n_sum + mo_sum + p_sum + q_sum + r_sum
            if sentaku2_sum >= 17:
                sentaku2_lst = ["",'選択（専門）',17,11,17]
                sentaku2_sum = sentaku2_sum - 17
                m_count = m_count + 1
                m_m_count = m_m_count + 1
            elif (sentaku2_sum >= 11) & (sentaku2_sum < 17):
                sentaku2_lst = ["",'選択（専門）',17,11,sentaku2_sum]
                sentaku2_sum = 0
                m_m_count = m_m_count + 1
            else:
                sentaku2_lst = ["",'選択（専門）',17,11,sentaku2_sum]
                sentaku2_sum = 0

            #学部専門系科目又は教養系科目
            sonota_sum = sentaku_sum + sentaku2_sum
            sonota_lst = ["",'学部専門系科目又は教養系科目',6,6,sonota_sum]
            o_count = 0 #卒業
            o_o_count = 0  #進級

            if sonota_sum >= 6:
                o_count = o_count + 1
                o_o_count = o_o_count + 1

            #CSV破損検知
            if all1_sum != tanni_sum or specialtanni > 0:
                error_message = "CSVファイルが破損,もしくは特殊な単位があるかもしれません"
                error_message_label.config(text=error_message)
                return

            #判定
            all_sum = n_count + m_count + o_count #卒業
            all_all_sum = n_n_count + m_m_count + o_o_count   #進級
            if all_all_sum == 14:
                result_text +='\n\n4年次進級判定：進級できます!'
            else:
                result_text +='\n\n4年次進級判定：進級できません!'

            if all_sum == 14:
                result_text +='\n\n卒業判定：卒業できます!\n-------------------'
            else:
                result_text +='\n\n卒業判定：卒業できません!\n-------------------'

            #進級
            if n_n_count == 9:
                result_text += '\n\n教養科目の進級要件：満たしている'
            else:
                result_text += '\n\n教養科目の進級要件：満たしていない'

            if m_m_count == 4:
                result_text += '\n\n学部専門系科目の進級要件：満たしている'
            else:
                result_text += '\n\n学部専門系科目の進級要件：満たしていない'

            if sonota_sum >= 6:
                result_text += '\n\n学部専門系科目又は教養系科目の進級要件：満たしている\n-------------------'
            else:
                result_text += '\n\n学部専門系科目又は教養系科目の進級要件：満たしていない\n-------------------'

            
            #卒業 
            if n_count == 9:
                result_text += '\n\n教養科目の卒業要件：満たしている'
            else:
                result_text += '\n\n教養科目の卒業要件：満たしていない'

            if m_count == 4:
                result_text += '\n\n学部専門系科目の卒業要件：満たしている'
            else:
                result_text += '\n\n学部専門系科目の卒業要件：満たしていない'

            if sonota_sum >= 6:
                result_text += '\n\n学部専門系科目又は教養系科目の卒業要件：満たしている'
            else:
                result_text += '\n\n学部専門系科目又は教養系科目の卒業要件：満たしていない'
            
            #表   
            table_data = [a_lst,b_lst,c_lst,d_lst,f_lst,i_lst,fgh_lst,ij_lst,sentaku_lst,m_lst,n_lst,mo_lst,sentaku2_lst,sonota_lst]

            for item in tree.get_children():
                tree.delete(item)
            
            #単位数が届いていなかった場合、その欄をハイライトする
            for data_row in table_data:
                if data_row[4] < data_row[3]:
                    tree.insert('','end',values=(data_row),tags='red,light_green')
                elif data_row[4] < data_row[2]:
                    tree.insert('','end',values=(data_row),tags='red,yellow')
                else:
                    tree.insert('','end',values=(data_row))
            tree.tag_configure('red,yellow',foreground='red')#文字の色を赤に変える
            tree.tag_configure('red,yellow',background='yellow')#背景を黄色に変える
            tree.tag_configure('red,light_green',foreground='red')
            tree.tag_configure('red,light_green',background='light green')

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
root.geometry("1000x750") #初期ウィンドウサイズ
root.title("令和５年度入学 工学部 知能情報システム ４年次進級＆卒業判定アプリ")

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


#表
style = ttk.Style()
style.configure("Treeview.Heading", font=('Yu Gothic', 12)) #列見出しのフォント
style.configure("Treeview", font=('Meiryo', 10)) #表のフォント

tree = ttk.Treeview(root, columns=("Col1", "Col2", "Col3", "Col4", "Col5"))
tree.heading("#1", text="区分")
tree.heading("#2", text="細区分")
tree.heading("#3", text="卒業要件")
tree.heading("#4", text="4年次進級要件")
tree.heading("#5", text="あなたの単位数")
tree.column('Col1', width=100)
tree.column('Col2', width=300)
tree.column('Col3', width=100)
tree.column('Col4', width=140)
tree.column('Col5', width=150)
tree.pack()
tree["show"] = "headings" #左１列表示させない
tree.configure(style="Treeview") #スタイルの適用

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
