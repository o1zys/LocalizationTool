from tkinter import *
import tkinter.messagebox
import main
import error
import global_var as gl


def get_latest():
    # set_path()
    #if version.get() == "":
    #    hintLabel["text"] = "必须填写Version! 请按照当前版本填，如0.6.0"
    #    hintLabel["fg"] = "red"
    #    return

    global done_flag
    if done_flag:
        hintLabel["text"] = ""
        a = tkinter.messagebox.askokcancel('警告', '您已成功生成了一次表格，再次生成可能会覆盖修改的颜色提示信息，确认要生成吗？')
        if not a:
            return

    error.Error.set_code(-1, "")
    hintLabel["text"] = ""
    main.execute(gl.csv_dir, gl.index_file, gl.trans_file, version.get())
    error_code = error.Error.get_code()
    if error_code == -1:
        hintLabel["text"] = "Success!"
        hintLabel["fg"] = "green"
        done_flag = True
    else:
        hintLabel["text"] = "[Error " + str(error_code) + "] " + error.Error.get_info(error_code)
        hintLabel["fg"] = "red"


def gen_task():
    if version.get() == "":
        hintLabel["text"] = "必须填写Version! 请按照当前版本填，如0.6.0"
        hintLabel["fg"] = "red"
        return

    error.Error.set_code(-1, "")
    hintLabel["text"] = ""
    main.gen_task(gl.trans_file, gl.task_file, version.get())
    error_code = error.Error.get_code()
    if error_code == -1:
        hintLabel["text"] = "Success!"
        hintLabel["fg"] = "green"
    else:
        hintLabel["text"] = "[Error " + str(error_code) + "] " + error.Error.get_info(error_code)
        hintLabel["fg"] = "red"


def update_glossary():
    print("update_glossary")
    error.Error.set_code(-1, "")
    hintLabel["text"] = ""
    main.update_glossary(gl.trans_file, gl.glossary_file, version.get())
    error_code = error.Error.get_code()
    if error_code == -1:
        hintLabel["text"] = "Success!"
        hintLabel["fg"] = "green"
    else:
        hintLabel["text"] = "[Error " + str(error_code) + "] " + error.Error.get_info(error_code)
        hintLabel["fg"] = "red"


def export_csv():
    print("export_csv")
    error.Error.set_code(-1, "")
    hintLabel["text"] = ""
    main.export_csv(gl.trans_file, gl.glossary_file)
    error_code = error.Error.get_code()
    if error_code == -1:
        hintLabel["text"] = "Success!"
        hintLabel["fg"] = "green"
    else:
        hintLabel["text"] = "[Error " + str(error_code) + "] " + error.Error.get_info(error_code)
        hintLabel["fg"] = "red"


#def convert():
#    set_path()
#    error.Error.set_code(-1, "")
#    hintLabel["text"] = ""
#    main.convert(lua_dir.get(), output_file.get())
#    error_code = error.Error.get_code()
#    if error_code == -1:
#        hintLabel["text"] = "Success!"
#        hintLabel["fg"] = "green"
#        fp = open("config2.txt", 'w')
#        fp.write(lua_dir.get()+"\n")
#    else:
#        hintLabel["text"] = "[Error " + str(error_code) + "] " + error.Error.get_info(error_code)
#        hintLabel["fg"] = "red"


root = Tk()
root.title('Localization Tool')
root.columnconfigure(1, weight=1)
# set global variable
try:
    gl.set_var_from_config()
except Exception as e:
    tkinter.messagebox.showerror('错误', '读取配置文件config.txt失败\n' + str(e))

global done_flag
done_flag = False
# root.iconbitmap(".\\icon.ico")

version = StringVar()
Label(root, text="Version:").grid(row=1, column=0)
Entry(root, textvariable=version).grid(row=1, column=1, sticky="nsew")

Button(root, text="Get Latest", command=get_latest).grid(row=2, columnspan=3)
Button(root, text="Generate Task", command=gen_task).grid(row=3, columnspan=4)
Label(root, text="").grid(row=4, column=4)
Button(root, text="Update Glossary", command=update_glossary).grid(row=5, columnspan=4)
Button(root, text="Export CSV", command=export_csv).grid(row=6, columnspan=4)

#Label(root, text="Lua Dir:").grid(row=6, column=0)
#Entry(root, textvariable=lua_dir).grid(row=6, column=1, sticky="nsew")
#Button(root, text="Select", command=select_lua_dir).grid(row=6, column=2)
#Button(root, text="Convert to Lua", command=convert).grid(row=7, columnspan=3)

hintLabel = Label(root, text="")
hintLabel.grid(row=8, columnspan=3)
Label(root, text="           -- Presented by Oizys").grid(row=9, columnspan=3, sticky=SE)
root.mainloop()

