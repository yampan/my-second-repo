"""
Font List Sample

repo: https://github.com/yampan/my-first-repo.git

git 操作は、Pi5.local で行う。
pi5a.local:/home/jupyter/work/GUI-test/my-first-repo/fontlist.py

  作業ディレクトリ(Working Directory)、索引(Index)、コミット、
  作業最後のコミットを指す HEAD
  
  1.共有リポジトリがない場合、リポジトリを作成
   git init
  
  2.共有リポジトリを、クローン(clone)して作業ディレクトリを作成
    git clone https://username@domain/path/to/repository
    
  3. ファイルの追加 & コミット
    git add <filename>, or git add *
  
  4. git commit -m "1st commit"
    変更内容が索引からコミットされ、HEADに格納されました。
  5. 共有リポジトリにプッシュする
    git push origin master
  
  6. ローカルでリポジトリを作成( git init )や共有リポジトリからクローン(clone)していない場合、共有リポジトリを登録することができます。
    git remote add origin <server>
  
  7. 作業ディレクトリを最新のコミットに更新
    git pull
      

ref: get_clipboard(), set_clipboard(), screenshot(),
     load_json_file(fn), save_json_file(fn, dat)

"""
import TkEasyGUI as eg
import json

# 定数
f_size = 14 # 16
fn_font = "fontlist0.json"


# list fonts
font_items2 = sorted(eg.get_font_list())
font_items=[]
f_dic = {}
if 1: # read  encoding="utf-8-sig" はBOM対策
  with open(fn_font, "r", encoding="utf-8-sig") as f:
      f_dic = json.load(f)
  print("{fn_font} is opened.")
  for i in f_dic.keys():
    if f_dic[i]:
      font_items.append(i)
  font_items = sorted(font_items)
else:
  for f in font_items2:
      if f[0:1] in ['@', "$", "%", "&"]: continue
      elif "HGrep" in f or "CR" in f or "AR" in f or "HG" in f or \
          "Jsut" in f or "ＤＦ" in f or "明朝" in f or "TA" in f or \
            "symbol" in f.lower():
          continue
      font_items.append(f)
      f_dic[f] = 1
# save font-name
#with open(fn_font, "w") as f:
#    json.dump(f_dic, f, indent=2, ensure_ascii=False)
#print("fontlist0.json was created.")

# define layout
col1 = eg.Column([[eg.Listbox(values=font_items, size=(40, 20),
            key="-fontlist-", enable_events=True, ),] ] )
col2 = eg.Column([ [eg.Button("enable", )],
                  [eg.Button("disable")], ], expand_x=True )
layout = [
    [
        eg.Frame(
            "Sample",
            expand_x=True,
            layout=[[eg.Text(f"Hello, 123 こんにちは?  Size:{f_size}", key="-sample-")]],
        )
    ],
    [col1, col2],
    [eg.Input("-", key="-font-", expand_x=True), eg.Button("Copy")],
    [eg.Button("font +"), eg.Button("font -"), eg.Text("   "),
     eg.Button("Save", color="#2222FF", font=("Arial",16,"bold")),eg.Text("  "),
     eg.Button("Exit", color="#FF2222", font=("Arial",16,"bold"))],
]
# create Window
flag = 1
with eg.Window("Font List", layout, font=("Arial", f_size), finalize=True, 
               resizable=True, center_window=False, location=(100,100)) as window:
    if flag:
        flag = 0

        print("get_size=", window.get_size())        
        print("get_center_location=", window.get_center_location())
        print("get_screen_size=", window.get_screen_size())
        aaa = 0.95
        print("set_alpha_channel=", aaa)
        window.set_alpha_channel(aaa)
        w_size = (600,700)
        print("set_size=", w_size)
        window.set_size(w_size)
    # set_theme
    eg.set_theme("clam")
    ''' - macOS --- ('aqua', 'clam', 'alt', 'default', 'classic')
    - Windows --- ('winnative', 'clam', 'alt', 'default', 'classic', 'vista', 'xpnative')
    - Linux --- ('clam', 'alt', 'default', 'classic')'''
    # event loop
    for event, values in window.event_iter():
        print("# event:", event, values)
        if event == "Exit" or event == eg.WINDOW_CLOSED:
            break
        if event in ["-fontlist-", "font +", "font -"]:
            if event == "font +": f_size += 1
            if event == "font -": f_size -= 1
            print("f_size=", f_size)
            # get font name from listbox
            fontlist: eg.Listbox = window["-fontlist-"]
            index = fontlist.get_cursor_index()
            if index >= 0:
                font_name = font_items[index]
                window["-font-"].update(f"font_name   :{f_dic[font_name]}")
                #window["-sample-"].update(font=(font_name, 18))
                window["-sample-"].set_text(f"Hello, 123 こんにちは?  Size:{f_size}")
                window["-sample-"].update(font=(font_name, f_size))
        if event == "Copy":
            eg.set_clipboard(values["-font-"])
            eg.print("Copied to clipboard:\n" + values["-font-"])
        if event == "Save":
            with open(fn_font, "w", encoding="utf-8") as f:
                json.dump(f_dic, f, indent=2, ensure_ascii=False)
            print(f"fontlist is saved in '{fn_font}'")
        if event == "enable":
            f_dic[font_name] = 1
            print("f{font_name} is set to 1")
            window["-font-"].update(f"font_name   :{f_dic[font_name]}")
        if event == "disable":
            f_dic[font_name] = 0
            print("f{font_name} is set to 0")
            window["-font-"].update(f"font_name   :{f_dic[font_name]}")
    # --- for loop end.
# --- END
print("Normal end.")


