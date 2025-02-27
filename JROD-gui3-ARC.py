### JROD-gui3-ARC (Branch: search_ando_sort)
### MS VS-Code Himt
# 1. ターミナルを開く　　ctrl + @, clear: cls,  Ctrl+Shift+P
# 2. GitHUB commitの方法
# 3. module のSub-folder を mylib に変更。2025-02-21
# 4. main へ subscribe する。
# 5. Branch logfunc を追加する。
# 6. acubic-PE d:\python_test\TkEasyGUI\my-second-repo\
#    PRO4: document\GITHUB\
"""
Font List Sample

TkEasyGUI
ref: https://github.com/kujirahand/tkeasygui-python
     

Pro4:  TkEasyGUI-test を pip install したが、import でエラーとなる。
       Python3.12 を再インストールする。
acubic-PE:
       D:/python_test/TkEasyGUI/my-first-repo
       
repo: https://github.com/yampan/my-first-repo.git

git 操作は、Pi5a.local で行う。(V:)
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
  
  6. ローカルでリポジトリを作成( git init )や共有リポジトリから
     クローン(clone)していない場合、共有リポジトリを登録することができます。
    git remote add origin <server>
  
  7. 作業ディレクトリを最新のコミットに更新
    git pull
      
=============================================================================
JROD-gui.py:
    1) GUI用に、Fontを選択する。
    2)

ref: get_clipboard(), set_clipboard(), screenshot(),
     load_json_file(fn), save_json_file(fn, dat)

"""

import TkEasyGUI as eg
import json, os, sys, datetime
import pytz
from mylib.readWriteXL import (openXl, getRow, setRow, search, JST, createLogWs,
                               j_map, trans2, logHeaderSet, JSTfn )
from mylib.db_access import query, DBtrans, db_init
from mylib.logger import (FMT, FMT2, createLogger, 
            clearLogfile, log_init, get_file_info)
import glob
from mylib.version_info2 import help_window
from mylib.daycheck import dayCheckM, pred

# logger start
os.makedirs("./log", exist_ok=True)
LOG_FN = "LOG_JROD.TXT"
DB_CONNECT = False

logger = log_init(LOG_FN)
get_file_info(".", "LOG*", show=1)

script_path = os.path.abspath(sys.argv[0])
script_name = os.path.basename(script_path)
logger.debug(f"=== Start: {script_name} ===")
logger.debug(f"スクリプトのパス: {script_path}")
logger.debug(f"スクリプト名: {script_name}")

logger.debug("db_init() = {db_init()}")
if DB_CONNECT:
    pass
else:
    logger.debug(f'DB_CONNECT: {DB_CONNECT}, skip DB.')
    
fn_conf = "JROD_config.json"
with open(fn_conf, "r") as f:
    f_dic = json.load(f)

# font
font_items = list(f_dic.keys())
f_size = f_dic["f_size"]
sel_font = f_dic["sel_font"]

# 定数
PTR = 2 # 1: header, 2: 実際のデータ。
FN_EXCEL = 'JRODe_ARC_Sample.xlsx'
SHEET_NAME = 'Sheet1'

# test data
id = "12345"
kannri_id = "kan-123"
name = "test patient"
sex = "M"
disease = "Lung ca."
dis_icdo = "C40.1"
pathology = "Adeno ca."
path_icdo = "M8140/3"

st_date = "2020-01-01"
en_date = "2020-01-20"
frac = "30"
dose = 60
days = 61
low = 30.5
high = 45.12345
comp = '完遂'
comp_pre = '中断あり'
status, final_d = 'death', '2020-01-01'
status2, final_d2 = 'dead', '2022-12-31'

#項目選択
comps =["予定治療完遂","予定治療完遂(8日以上の中断あり)","予定の50%未満で中止","予定の50%以上で中止",
        "遂行程度不詳で中止","その他","不明"]
stats = ['1.非担癌生存','2.担癌生存','3.担癌不詳生存','4.原病死','5.他病死','6.不明死','7.消息不明']

status_ARC = {'13111':'1.非担癌生存','13114':'4.原病死', '13113':'3.担癌不詳生存', 
              '13112':'2.担癌生存','13115':'5.他病死','13116':'6.不明死', '13117':'7.消息不明' }


# excel open  wb:workbook, ws:worksheet, title:dict {}
logger.debug(f"fn = {FN_EXCEL}, sheet = {SHEET_NAME}")
wb, ws, title = openXl(FN_EXCEL, SHEET_NAME)
ws2, ws3 = createLogWs(wb, ws)
'''
ws2 = wb.copy_worksheet(ws)
ws2.title = "ws_copy"
ws3 = wb.create_sheet(title = 'ws_log')
logHeaderSet(ws3)
'''
logger.debug(f"sheet names = {wb.sheetnames}")
logger.debug(f"ws2: 'copied-ws',  ws3: 'trans-log' were opened.")

# title: dict ==> title2: list
title2 = ['index']
for i in range(1,ws.max_column+1):
    title2.append(title[i])


# PTRにより、データの読み出し
def setByMap(j_map, ws, PTR, window, deb=0):
    global id, kannri_id, name, sex, disease, dis_icdo, pathology, path_icdo
    global st_date, en_date, frac, dose, days, low, high, comp, comp_pre
    global final_d, status

    col = getRow(ws, PTR)
    #logger.debug(f"col = {col}")
    if deb: print("col =", col)
    
    #    Global 変数へ展開
    for v in j_map.keys():
        (ptr, nam) = j_map[v] 
        #print(f"{v}: ptr:{ptr}, nam:{nam}")
        cmd = f"{v} = col[{ptr}]"
        if deb: print(f"setByMap:  {v}: ptr:{ptr}, nam:{nam},  cmd= '{cmd}'")
        exec(cmd,locals(),globals())
    logger.debug(f"id={id}, comp={comp}, status={status}," )
    
    # dayCheckM()
    NG, mes, (low, period, high) = dayCheckM(col, PTR, title2, sys.stdout)
    comp_pre = pred(col, low, period, high, '85:放射線治療完遂度')
    # DB read
    final_d2, status2 = DBread(kannri_id)
    logger.debug(f"final_d2={final_d2}, status2={status2}")
    
    # redraw
    window["-id-"].update(f"ID: {id:10}, ")
    window["-id2-"].update(f" kanri: {kannri_id:10}, name: {name:15}")
    window["-dis-"].update(f"{disease:15},({dis_icdo:5}) / {pathology},({path_icdo})")
    window["-date-"].update(f"開始日:{st_date}, 終了日:{en_date}  Dose:{dose}, Frac:{frac}, days:{days},")
    if int(days) < low or int(days) > high:
        color = "red"
    else:
        color = "black"
    window["-comp-"].update(f"{low:8} < {days:3} < {high:8.1f},     元の完遂度: {comp}",
                            color=color)
    window["-comp2-"].update(f"完遂予測：{comp_pre} ----->  ")
    window["-ptr-"].update(f"{PTR}")
    window["-final_d-"].update(f"{final_d}")
    window["-status0-"].update(f"  生死の状況: {status:6} ==> ")
    window["-status-"].update(f'{status}')
    window["-info-"].update(f"  {JST()}")
    window["-final_d2-"].update(f' DB: 　{final_d2}     生死の状況：{status2}')
    window["-s_error-"].update("")
    return


# PTRによるデータの書き出し
def returnByMap(j_map, ws, PTR, deb=0):
    global id, kannri_id, name, sex, disease, dis_icdo, pathology, path_icdo
    global st_date, en_date, frac, dose, days, low, high, comp, comp_pre
    global final_d, status
    
    col = getRow(ws, PTR)
    for v in j_map.keys():
        (ptr, nam) = j_map[v] 
        #print(f"{v}: ptr:{ptr}, nam:{nam}")
        cmd = f"col[{ptr}] = {v}"
        if deb: 
            logger.debug(f"returnByMap:  {v}: ptr:{ptr}, nam:{nam},  cmd= [{cmd}]")
        exec(cmd,locals(),globals())
    # -- "comp":(85,'放射線治療完遂度')   'status':(87,'生死の状況')
    logger.debug(f"returnByMap: #150 comp={comp}, col[85]={col[85]}, status={status}, col[87]={col[87]}")
    setRow(ws, PTR, col)
    return


# DBからデータの読み出し
def DBread(id):
    """
    DB から SQLを実行し、final_d2, status2 をセットする。
    Args:
        id: 管理番号
    Returns:
        (final_d2, status2) タプルで返す
    """
    if DB_CONNECT:
        sql = '''select pat_id1,user_defined_dttm_1,user_defined_pro_id_3 
                from admin where pat_id1 = ''' + f"'{id}' ;" 
        logger.debug(f'sql = {sql}')
        rows = query(sql)
        logger.debug(f'rows = {rows}')
    
        # None の処理
        if rows[0][1] is None or rows[0][2] is None:
            rows = [(16119, datetime.datetime(2023, 4, 29, 0, 0), 13114)]
    else:
        rows = [(16119, datetime.datetime(2023, 4, 29, 0, 0), 13114)]
    
    # value set
    final_d2 = rows[0][1]
    if type(final_d2) is not str:
        final_d2 = f'{final_d2}'[:10]
    status2 = status_ARC[f'{rows[0][2]}']
        
    return (final_d2, status2)


# DBへデータの書き込み
def DBwrite(id, dt, st):
    """
    DB に SQLを実行し、final_d2, status2 をセットする。
    Args:
        id: where 管理番号 = id
        dt: 最終確認日
        st: 病態
    Returns:
        None
    """
    status2 = None
    for k,v in status_ARC.items():
        if st in v:
            status2 = int(k)
    if status2 is None:
        print(f'status ERROR: {st} is not found.')
        logger.debug(f'status ERROR: {st} is not found.')
        return
    values = (status2, dt, id)
    sql = '''update admin set user_defined_pro_id_3 = ? , 
                user_defined_dttm_1 = ?  
                where pat_id1 = ? ;'''
    #print(f'#DBwrite: sql: {sql},\n  values: {values}')
    logger.debug(f'#DBwrite: sql: {sql},\n  values: {values}')
    if DB_CONNECT:
        DBtrans(sql, values)
    return 





### eg.Text("click me", font=("Arial", 30,'bold italic'), enable_events=True, 
#            background_color="red", text_color="white"),

# define layout
lay_info=[[eg.Text(f"sel_font: {sel_font},  Size:{f_size},", font=("Arial",12,"bold"), 
                background_color="lightyellow", key="-sample-"),
           eg.Text(" ", background_color="lightyellow", expand_x=True),
           eg.Text(f"  {JST()}", font=("Arial",12,"bold italic"), color="green", 
                background_color="lightyellow", key="-info-")],
          [eg.Text(f"file: {FN_EXCEL}, sheet: {SHEET_NAME},  max_row:{title['max_row']}"+ \
                   f",   max_col:{title['max_column']}", font=("Arial",12,'bold'),
                   background_color="lightyellow",), 
           eg.Text(" ", expand_x=True, background_color="lightyellow",),
           eg.Button("HELP", font=("Arial", 10, 'bold')), ],
         ]
lay_status = [
    [eg.Input(f"{final_d}", width=12, background_color="lightyellow", key="-final_d-"),
     eg.Text(f"  生死の状況: {status:6} ==> ", 
             background_color="lightyellow", key="-status0-"),
     eg.Input(f'{status}', width=12, key="-status-"), 
     eg.Button("fix2", font=("Arial",13,'bold'), color="purple",
               background_color="lightyellow"), ],
    [eg.Text(f' DB: 　{final_d2}     生死の状況：{status2}', 
             font=("BIZ UDPゴシック", 12, "bold"),
             color='blue', key="-final_d2-"), eg.Text("    ", expand_x=True),
     eg.Checkbox("DBconnect ", key="-DBconnect-", enable_events=True),],
    ]
lay_tnm = [
    [eg.Text(" cpr"), eg.Input("c", width=2, key="-cpr-"),
     eg.Text(" T"), eg.Input("T1", width=2, key="-T1-"),
     eg.Text(" N"), eg.Input("N1", width=2, key="-N1-"),
     eg.Text(" M"), eg.Input("M1", width=2, key="-M1-"),
     eg.Text("  Stage"), eg.Input("1a", width=4, key="-S1-"),
     eg.Text("     ", ), 
     eg.Button("fix3",font=("Arial",13,'bold'), color="purple",
               background_color="lightyellow"), ],    
]
layout = [
    [eg.Frame(f" JROD-GUI2 Project: {script_name}  TkEasyGUI ver: {eg.__version__} ", expand_x=True,
            layout=lay_info, font=("Arial",10,'bold'), background_color="lightyellow",color="blue") ],
    [eg.Text("  ",font=("Arial",5,'bold'),),],
    [eg.Text(f"ID: {id:10}, ", key="-id-"),
     eg.Text(f" kanri: {kannri_id:10}, name: {name:15}", key="-id2-"), ],
    [eg.Text(f"{disease:15},({dis_icdo:5}) / {pathology},({path_icdo})", key="-dis-")],
    [eg.Input("管理番号", width=10, key="-kanri_no-"),
     eg.Button("search"), eg.Text("", width=10, key="-s_error-", color="red"),
     eg.Text("---------------------- ID ==> ", ), 
     eg.Button("paste", font=("Arial",13,'bold'), color="purple",),],
    [eg.Text(f"開始日:{st_date}, 終了日:{en_date}  Dose:{dose}, Frac:{frac}, days:{days},", key="-date-")],
    [eg.Text(f"{low:8} < {days} < {high:8.2f},     元の完遂度: {comp}", key="-comp-")],
    [eg.Text(f"      完遂予測:{comp_pre} ----->  ", key="-comp2-"), 
     eg.Input("---", key="-font-", width=22,), 
     eg.Button("fix", font=("Arial",13,'bold'), color="purple",background_color="lightyellow"),],
    [eg.Listbox(values=comps, size=(22, 7), key="-complist-", enable_events=True, ),eg.Text("   ↑     "),
     eg.Listbox(values=stats, size=(10,7), key="-statlist-", enable_events=True, ), eg.Text("  ↓") ],
    [eg.Frame(" 最終確認日 ", font=("Arial", 12, 'bold'), expand_x=True, layout=lay_status, ),],
    [eg.Frame("TNM Stage", layout=lay_tnm, font=("Arial", 12, 'bold'), color="green", 
              expand_x=True)],
    #
    #[eg.Text("-----------------------------------------------------------", ),],
    [eg.Text("PTR: "), eg.Input(f"{PTR}", key="-ptr-", enable_events=False, width=5,),
     eg.Button("set"), eg.Text("    "), 
     eg.Button("< prev"), eg.Button("next >"),
     eg.Text("   　　　"),
     eg.Button("Save", color="#2222A0",font=("Arial",14,"bold")),eg.Text("   "),
     eg.Button("Exit", color="#FF2222", font=("Arial",14,"bold")),
     eg.Text("     ", expand_x=True), 
     eg.Button("clear", font=("Arial",10,'bold'),color="brown",background_color="lightblue"), ],
    [eg.Multiline(text="message:", size=(40, 10), key="-body-",
            font=("Arial",11,'bold'), expand_y=True, expand_x=True)],
    [eg.Text(f' ', expand_x=True), eg.Text(f"JROD-gui2-ARC ver. 1.1", font=("Arial",11,'bold italic')) ]
]
# create Window
flag = 1 # メイリオ,"Arial"
with eg.Window(f"JROD-GUI: {script_name}", layout, font=(sel_font, f_size), finalize=True,
                 resizable=True, center_window=False, location=(10,10)) as window:
    if flag:
        flag = 0
        logger.debug(f"get_center_location= {window.get_center_location()}")
        logger.debug(f"get_screen_size= {window.get_screen_size()}")
        aaa = 0.98
        logger.debug(f"set_alpha_channel= {aaa}")
        window.set_alpha_channel(aaa)
        w_size = (700,900) # Width, Height
        logger.debug(f"set_size= {w_size}")
        window.set_size(w_size)
        logger.debug(f"get_size= {window.get_size()}")
        setByMap(j_map, ws, PTR, window)
        window["-body-"].print(f"\nfn = '{FN_EXCEL}', max_col = {ws.max_column}, max_row = {ws.max_row}",
                               text_color="purple")
    # event loop
    for event, values in window.event_iter(timeout=1000): # 1000 = 1 sec.
        if event == "-TIMEOUT-":
            window["-info-"].update(f"  {JST()}")
            continue
        values.pop("-body-")
        print(f"# event: {event}, values: {values}")
        
        if event == "Exit" or event == eg.WINDOW_CLOSED:
            break
        if event == "Save":
            f_dic["PTR"] = PTR
            f_dic["FN_EXCEL"] = FN_EXCEL
            
            with open("fontlist.json", "w") as f:
              json.dump(f_dic, f, indent=2, ensure_ascii=False)
            logger.debug("#save to 'JRODe_test.xlsx'")
            returnByMap(j_map, ws, PTR)
            fn_excel = f"JRODe_{JSTfn()}.xlsx"
            wb.save(fn_excel)
            window["-body-"].print(f"saved to ('fontlist.json', '{fn_excel}')",
                                   text_color="purple")
        if event == "-statlist-":
            statlist: eg.Listbox = window["-statlist-"]
            index = statlist.get_cursor_index()
            if index >= 0:
                status = stats[index]
            #val = values["-statlist-"]
            print(f"status = {status}")
            window["-status-"].update(status)
        if event in ["-complist-"]:
            complist: eg.Listbox = window["-complist-"]
            index = complist.get_cursor_index()
            if index >= 0:
                comp = comps[index]
            print("comp=", comp, type(comp))
            window["-font-"].update(comp)
        if event in ["fix", "fix2"]:
            final_d = values["-final_d-"]
            logger.debug(f"comp= {comp}, status= {status}, final_d= {final_d}")
            comp = values["-font-"]
            status = values["-status-"]
            logger.debug(f"comp= {comp}, status= {status}, final_d= {final_d}")
            window["-comp-"].update(f"{low:8} < {days} < {high:8.2f},    data:{comp}")
            window["-status-"].update(f"{status}")
            window["-status0-"].update(f"  0生死の状況: {status} ==> ")
            if event == "fix":
                logger.debug(f"trans2 = {trans2(ws2, PTR, {'comp':comp}, ws3)}")
                logger.debug(f"trans2 = {trans2(ws2, PTR, {'status':status}, ws3)}")
            if event == 'fix': window["-body-"].print(event, end=", ", text_color="purple")
            returnByMap(j_map, ws, PTR)
            if event == 'fix2':
                DBwrite(kannri_id, final_d, status)
                window["-body-"].print(event, end=", ", text_color="purple")
                logger.debug(f"trans2 = {trans2(ws2, PTR, {'final_d':final_d}, ws3)}")
                logger.debug(f"trans2 = {trans2(ws2, PTR, {'status':status}, ws3)}")
        if event in ["-ptr-", "< prev", "next >", "set"]:
            if event == "< prev" and PTR >2: PTR -= 1
            if event == "next >" and PTR < ws.max_row: PTR += 1
            if event == "set": PTR = int(values["-ptr-"])
            #print("PTR=", PTR)
            window["-ptr-"].update(f"{PTR}")
            setByMap(j_map, ws, PTR, window)
        if event == "paste":
            eg.set_clipboard(id)
            eg.print("Copied to clipboard:\n" + f"[{id}]" )
        if event == "clear":
            window["-body-"].update("=== cleared. ===")
        if event == "HELP":
            win = help_window(script_name)
            win.close()
        if event == "-DBconnect-":
            DB_CONNECT = values["-DBconnect-"]
            window["-body-"].print(f'DB_CONNECT = {DB_CONNECT}')
        if event == "search":
            val = values["-kanri_no-"]
            logger.debug(f"search: (ws, 1, {val}, 2)")
            scol, sPTR = search(ws, 1, val, 2)
            if sPTR is not None:
                logger.debug(f"scol={scol[:10]} ..., sPTR={sPTR}")
                PTR = sPTR
                setByMap(j_map, ws, PTR, window)
            else:
                logger.debug(f"search: NOT FOUND, sPTR={sPTR}")
                window['-s_error-'].update("Not found.")
            window["-body-"].print(f"search: {val}, sPTR={sPTR}", text_color="purple")
        # LOG
        text = f"#event:{event}, PTR:{PTR}, comp:{comp}, final_d:{final_d}, status:{status}"
        
        window["-body-"].print(text, text_color="darkblue", ) # background_color="lightpink"
        #window["-body-"].update(text)
# ---
print("   ==> NORMAL END.")
