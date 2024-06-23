import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import *
import excel
from random import sample,shuffle
from tkinter.ttk import Treeview
from datetime import datetime


def windows3(username):
    def treeviewclick(event,tree):
        window3.clipboard_clear()
        strs=""
        for item in tree.selection():
            item_text=tree.item(item,"values")
            strs+=item_text[0]+"\n"					#获取本行的第一列的数据
        window3.clipboard_append(strs)

    window3=Tk()
    window3.geometry('550x250')
    cols = ("序号", "日期","种类","题数","成绩")
    ybar=Scrollbar(window3,orient='vertical')      #竖直滚动条
    tree=Treeview(window3,show='headings',columns=cols,yscrollcommand=ybar.set)
    ybar['command']=tree.yview
    #表头设置
    for col in cols:
        tree.heading(col,text=col)             #行标题
        tree.column(col,width=110,anchor='center')
    #插入数据
    data=excel.getData(username+"record.xls",sheetId=0,lis=["序号","日期","种类","题数","成绩"])
    for i in range(len(data)):
        tree.insert("","end",values=(int(data[i][0]),data[i][1],data[i][2],int(data[i][3]),data[i][4]))
    tree.grid(row=0,column=0)
    ybar.grid(row=0,column=1,sticky='ns')
    window3.mainloop()
def windows2(username):
    window.destroy()
    window2 = tk.Tk()  # 建立第二个窗口
    window2.title("Choose")
    window2.geometry("500x500")

    def b1():
        nonlocal username
        windows3(username)
    def b2():
        kind="四级"
        number=0
        scoreswd = tk.Toplevel(window2)
        scoreswd.geometry("500x200")
        scoreswd.title('Choose Word Practice')

        def func1(self):
            nonlocal kind
            kind=combobox1.get()

        def func2(self):
            nonlocal number
            number=int(combobox2.get())
        def yes():
            print(kind,number)
            data=excel.getData(kind+".xls",sheetId=0,lis=["单词","释义"],requireHeader=False,requestWholeTable=False)
            t=sample(data,number)
            scoreswd.destroy()
            score=0
            def check_answer():
                # 假设正确答案是"Option 3"
                nonlocal answer
                nonlocal score
                correct_answer =str(answer.get())

                selected_option = var.get()
                print(type(correct_answer)," ",type(selected_option))
                print(correct_answer," ",selected_option)
                print("correct:",correct_answer)
                print("selected:",selected_option)
                if selected_option == correct_answer:
                    messagebox.showinfo("Correct!", "You have selected the correct answer!")
                    score+=1
                else:
                    messagebox.showerror("Incorrect", "You have selected the incorrect answer.")
                ck.destroy()
            window2.destroy()
            for i in range(number):
                option=sample(data,3)
                while t[i] in option:
                    option=sample(data,3)
                option.append(t[i])
                print(option)
                shuffle(option)
                # 创建主窗口
                ck = tk.Tk()
                ck.geometry('400x400')
                ck.title("Single Choice Question")
                # 创建问题标签
                question_label = tk.Label( ck, text=t[i][0])
                question_label.pack(pady=20)

                # 创建变量来存储选中的选项
                var = tk.StringVar()
                answer=tk.StringVar()
                answer.set(t[i][1])
                print(t[i])

                # 添加单选按钮
                for j in range(4):
                    tk.Radiobutton(ck, text=option[j][1], variable=var, value=option[j][1]).pack(pady=10)

                # 添加提交按钮
                submit_button = tk.Button(ck, text="Submit", font=('楷体',12),fg='green',command=check_answer)
                submit_button.pack(pady=20)
                # 进入主循环
                ck.mainloop()

            data2=excel.getData(username+"record.xls",sheetId=0,lis=["序号","日期","种类","题数","成绩"],requireHeader=True,requestWholeTable=False)
            data2.append([data2[-1][0]+1,str(datetime.now()),kind,number,round(score/number,2)*100])
            excel.printToFile(username+"record",data2)
            messagebox.showinfo("成功","恭喜您答题成功，您的分数是"+str(round(score/number,2)*100))


        combobox1 = ttk.Combobox(scoreswd,font=('楷体', 12),width=15)
        combobox1.place(x=20, y=80)
        combobox1['values'] = ('四级', '六级', '课本', '雅思')
        #combobox1.current(0)
        combobox1.bind('<<ComboboxSelected>>', func1)
        combobox2 = ttk.Combobox(scoreswd,font=('楷体', 12),width=15)
        combobox2.place(x=200, y=80)
        combobox2['values'] = ('10', '20', '30')
        #combobox2.current(0)
        combobox2.bind('<<ComboboxSelected>>', func2)
        bu = tk.Button(scoreswd, text="确定", command=yes,font=('楷体', 14))
        bu.place(x=400, y=80)
        lb = tk.Label(scoreswd,text="请选择想练习的单元和单词个数：",font=('楷体', 18))
        lb.place(x=50, y=20)
        scoreswd.mainloop()

    def b3():
        def windows4(filename):
            def treeviewclick(event,tree):
                window4.clipboard_clear()
                strs=""
                for item in tree.selection():
                    item_text=tree.item(item,"values")
                    strs+=item_text[0]+"\n"					#获取本行的第一列的数据
                window4.clipboard_append(strs)
            window4=Toplevel(scoreswd)
            window4.geometry('520x250')
            cols = ("单词","释义")
            ybar=Scrollbar(window4,orient='vertical')      #竖直滚动条
            tree=Treeview(window4,show='headings',columns=cols,yscrollcommand=ybar.set)
            ybar['command']=tree.yview
            #表头设置
            tree.heading("单词",text="单词")             #行标题
            tree.column("单词",width=100,anchor='w')
            tree.heading("释义",text="释义")             #行标题
            tree.column("释义",width=300,anchor='w')
            #插入数据
            data=excel.getData(filename+".xls",sheetId=0,lis=["单词","释义"],requireHeader=False,requestWholeTable=False)
            for i in range(len(data)):
                tree.insert("","end",values=(data[i][0],data[i][1]))
            tree.grid(row=0,column=0)
            ybar.grid(row=0,column=1,sticky='ns')
            def addword():
                def yes1():
                    nonlocal filename
                    data=excel.getData(filename+".xls",sheetId=0,lis=["单词","释义"],requireHeader=True,requestWholeTable=False)
                    data.append([word.get(),mean.get()])
                    excel.printToFile(filename,data)
                    messagebox.showinfo("添加成功","添加成功！")
                addtk = tk.Toplevel(window4)
                addtk.geometry('350x200')
                addtk.title('Sign up window')

                word = tk.StringVar()
                tk.Label(addtk, text='英文: ').place(x=10, y= 10)
                entry_word= tk.Entry(addtk, textvariable=word)
                entry_word.place(x=150, y=10)

                mean = tk.StringVar()
                tk.Label(addtk, text='中文: ').place(x=10, y=50)
                entry_mean = tk.Entry(addtk, textvariable=mean)
                entry_mean.place(x=150, y=50)

                btn= tk.Button(addtk, text='确定', command=yes1)
                btn.place(x=150, y=130)


            def deleteword():
                def yes1():
                    nonlocal filename
                    data=excel.getData(filename+".xls",sheetId=0,lis=["单词","释义"],requireHeader=True,requestWholeTable=False)
                    for i in range(1,len(data)):
                        if data[i][0]==word.get():
                            data.pop(i)
                            break
                    else:
                        messagebox.showerror("删除失败！","英文不存在！")
                    excel.printToFile(filename,data)
                    messagebox.showinfo("删除成功","删除成功！")
                addtk = tk.Toplevel(window4)
                addtk.geometry('350x200')
                addtk.title('Sign up window')

                word = tk.StringVar()
                tk.Label(addtk, text='英文: ').place(x=10, y= 10)
                entry_word= tk.Entry(addtk, textvariable=word)
                entry_word.place(x=150, y=10)

                btn= tk.Button(addtk, text='确定', command=yes1)
                btn.place(x=150, y=130)
            def updateword():
                def yes1():
                    nonlocal filename
                    data=excel.getData(filename+".xls",sheetId=0,lis=["单词","释义"],requireHeader=True,requestWholeTable=False)
                    for i in range(1,len(data)):
                        if data[i][0]==word.get():
                            data[i][1]=mean.get()
                    excel.printToFile(filename,data)
                    messagebox.showinfo("修改成功","修改成功！")
                addtk = tk.Toplevel(window4)
                addtk.geometry('350x200')
                addtk.title('Sign up window')

                word = tk.StringVar()
                tk.Label(addtk, text='英文: ').place(x=10, y= 10)
                entry_word= tk.Entry(addtk, textvariable=word)
                entry_word.place(x=150, y=10)

                mean = tk.StringVar()
                tk.Label(addtk, text='中文: ').place(x=10, y=50)
                entry_mean = tk.Entry(addtk, textvariable=mean)
                entry_mean.place(x=150, y=50)

                btn= tk.Button(addtk, text='确定', command=yes1)
                btn.place(x=150, y=130)

            def F5():
                window4.destroy()
                windows4(combobox1.get())
            btn1 = tk.Button(window4, text="添加", command=addword,font=('楷体', 14))
            btn1.place(x=450,y=20)
            btn2 = tk.Button(window4, text="删除", command=deleteword,font=('楷体', 14))
            btn2.place(x=450,y=80)
            btn3 = tk.Button(window4, text="修改", command=updateword,font=('楷体', 14))
            btn3.place(x=450,y=140)
            btn4 = tk.Button(window4, text="刷新", command=F5,font=('楷体', 14))
            btn4.place(x=450,y=200)
            window4.mainloop()
        def func1(self):
            pass
        def yes():
            windows4(combobox1.get())
        scoreswd = tk.Toplevel(window2)
        scoreswd.geometry("300x300")
        scoreswd.title('Words')
        combobox1 = ttk.Combobox(scoreswd,font=('楷体', 12),width=15)
        combobox1.place(x=20, y=80)
        combobox1['values'] = ('四级', '六级', '课本', '雅思')
        combobox1.current(0)
        combobox1.bind('<<ComboboxSelected>>', func1)
        bu = tk.Button(scoreswd, text="确定", command=yes,font=('楷体', 14))
        bu.place(x=200, y=80)
        scoreswd.mainloop()


    canvas = tk.Canvas(window2, height=200, width=500)
    image_file = tk.PhotoImage(file='../word(1)/welcome3.gif')
    image = canvas.create_image(0,0, anchor='nw', image=image_file)
    canvas.place(x=25, y=350)

    tk.Label(window2, text='请选择你要进行的操作: ',font=('楷体',28,'bold')).place(x=40, y= 0)
    btn1 = tk.Button(window2, text='查看历史成绩',fg="orange",font=('楷体',20),command=b1)
    btn1.place(x=130, y=80)
    btn2 = tk.Button(window2, text='进行单词测试',fg="orange",font=('楷体',20),command=b2)
    btn2.place(x=130, y=170)
    btn3 = tk.Button(window2, text='修改单词库',fg="orange",font=('楷体',20),command=b3)
    btn3.place(x=130, y=260)
    window2.mainloop()
def usr_login():
    usr_name = var_usr_name.get()
    usr_pwd = var_usr_pwd.get()
    users=excel.getData("users.xls",sheetId=0,lis=["username","password"],requireHeader=True,requestWholeTable=False)
    for i in range(1,len(users)):
        if usr_name == users[i][0]:
            if usr_pwd == users[i][1]:
                tk.messagebox.showinfo(title='Welcome', message='How are you? ' + usr_name)
                windows2(usr_name)
            elif usr_pwd !=users[i][1]:
                tk.messagebox.showerror(message='Error, your password is wrong, try again.')
            break
    else:
        is_sign_up = tk.messagebox.askyesno('Welcome',
                               'You have not signed up yet. Sign up today?')
        if is_sign_up:
            usr_sign_up()

def usr_sign_up():
    def sign_to_Mofan_Python():
        np = new_pwd.get()
        npf = new_pwd_confirm.get()
        nn = new_name.get()
        users=excel.getData("users.xls",sheetId=0,lis=["username","password"],requireHeader=True,requestWholeTable=False)
        for i in range(1,len(users)):
            if np == users[i][0]:
                tk.messagebox.showerror('Error', 'The user has already signed up!')
                break
        else:
            if np != npf:
                tk.messagebox.showerror('Error', 'Password and confirm password must be the same!')
            else:
                users=excel.getData("users.xls",sheetId=0,lis=["username","password"],requireHeader=True,requestWholeTable=False)
                users.append([nn,np])
                tk.messagebox.showinfo('Welcome', 'You have successfully signed up!')
                window_sign_up.destroy()
        excel.printToFile(nn+"record",[["序号","日期","种类","题数","成绩"]])
        excel.printToFile("users",users)
    window_sign_up = tk.Toplevel(window)
    window_sign_up.geometry('350x200')
    window_sign_up.title('Sign up window')

    new_name = tk.StringVar()
    new_name.set('username')
    tk.Label(window_sign_up, text='User name: ').place(x=10, y= 10)
    entry_new_name = tk.Entry(window_sign_up, textvariable=new_name)
    entry_new_name.place(x=150, y=10)

    new_pwd = tk.StringVar()
    tk.Label(window_sign_up, text='Password: ').place(x=10, y=50)
    entry_usr_pwd = tk.Entry(window_sign_up, textvariable=new_pwd, show='*')
    entry_usr_pwd.place(x=150, y=50)

    new_pwd_confirm = tk.StringVar()
    tk.Label(window_sign_up, text='Confirm password: ').place(x=10, y= 90)
    entry_usr_pwd_confirm = tk.Entry(window_sign_up, textvariable=new_pwd_confirm, show='*')
    entry_usr_pwd_confirm.place(x=150, y=90)

    btn_comfirm_sign_up = tk.Button(window_sign_up, text='Sign up', command=sign_to_Mofan_Python)
    btn_comfirm_sign_up.place(x=150, y=130)
window = tk.Tk()
window.title('Welcome to Words Killer!')
window.geometry('450x300')

# welcome image
canvas = tk.Canvas(window, height=200, width=500)
image_file = tk.PhotoImage(file='welcome3.gif')
image = canvas.create_image(0,0, anchor='nw', image=image_file)
canvas.pack(side='top')

# user information
tk.Label(window, text='User name: ').place(x=50, y= 150)
tk.Label(window, text='Password: ').place(x=50, y= 190)

var_usr_name = tk.StringVar()
var_usr_name.set('username')
entry_usr_name = tk.Entry(window, textvariable=var_usr_name)
entry_usr_name.place(x=160, y=150)
var_usr_pwd = tk.StringVar()
entry_usr_pwd = tk.Entry(window, textvariable=var_usr_pwd, show='*')
entry_usr_pwd.place(x=160, y=190)
# login and sign up button
btn_login = tk.Button(window, text='Login', command=usr_login)
btn_login.place(x=170, y=230)
btn_sign_up = tk.Button(window, text='Sign up', command=usr_sign_up)
btn_sign_up.place(x=270, y=230)
window.mainloop()


