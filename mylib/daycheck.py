### 2025-02-14
# SRC から移植
import sys, math
import pandas as pd

# --- logger check
log = globals().get('logger', None)
if log is None:
    PRINT = print
else:
    PRINT = logger.debug 
print('logger =', log)

day_err = 0 # dayCheck でエラーになった件数。

def dayCheckCounter(n=None):
    global day_err
    
    if n is not None:
        day_err = n
    return day_err
 
        
def dayCheckM(col, n, title, f=sys.stdout, deb=0):
    global day_err
    '''
    開始日、終了日、回数、日数から完遂度をチェックする。
    args
        col: 1レコードのデータ(１から開始)  col[1],col[2], ... col[ws.max_column]
        n: pointer
        title: 項目名 list
        f: 出力先（ファイル）

    retrun err, s, (low, period, high)
        err: エラーの時は、１，OK=0
        s: 結果のメッセージ
        (low, period, high): low=予想最低、period=実際日数、high=予想最高
        
    # 043) 外部照射開始日　　　　:2022/11/18
    # 044) 外部照射終了日　　　　:2022/12/02
    # 045) 外部照射総線量　　　　:14.0
    # 046) 外部照射日数　　　　　:15.0
    # 047) 外部照射分割回数　　　:7.0
    # 048) 一日あたり照射回数　　:1.0
    '''
    #id = getID(str(col[1])[0:4]+"-"+str(col[1])[4:])
    id = col[112]
    if deb:
        print(f"\n  #{n:04} dayCheck: *** ptr:{col[0]} *** ID:{id} 管理番号：{col[1]}  {col[3]} {col[2]}", file=f)
    st, en = col[43], col[44]
    if st != st or en != en:
        print(f"     st={st}, en={en} *** ERROR ***",)
        s = f"     開始日st={st}, 終了日en={en} *** ERROR ***"
        err = 1
        return err, s
    # type check
    if deb:print(type(st), type)
    st = col[43]
    t = f"{type(st)}"
    if deb:print("#49 ", st, t)
    if "Timestamp" in t: 
        if deb:print("#51 timestamp")
        #st = ts2str(st)
    else:
        print(f"#55 st = {col[43]} type:{type(col[43])}")
        st = pd.Timestamp(col[43])
    if deb:print("#55 ", st, t)

    en = col[44]
    t = f"{type(en)}"
    if "Timestamp" in t: 
        if deb:print("#60 timestamp")
        #en = ts2str(en)
    else:
        en = pd.Timestamp(col[44])
    # --- st, en must dt
    if deb:print("#65  st=", type(st), " en=", type(en))
    
    period = (en - st).days + 1
    frac = int(col[47])
    frac0 = frac
    
    # Hyperの補正　48:一日あたり照射回数　２０２５－０２－０３
    if int(col[48]) > 1:
        frac = frac/int(col[48])
        print(f'     {title[48]}補正 = {frac0} ==> {frac}、　#{n:04}' , file=f)
        
    #print(f"st= {st}, en= {en}, period= {period}, frac= {frac}", file=f)
    # (最低日数）回数/5*7 + （マージン週に2日）回数/5*2
    base = round((frac/5*7 ), ndigits=3)
    # LOW: 最低日数 frac//5 * 7 に
    low = (frac-1)//5*7 + (frac%5 == 0)*5 + frac%5
    delta = round((frac/5*1.5+1), ndigits=3)
    high = round(base + delta, ndigits=3)
    
    # high は切り上げ ２０２５－０２－１０
    high0 = high
    high = math.ceil(high) 
    #if high > high0: print(f'     High:{high0} ==> {high}', file=f)

    if deb:print(f"base={base}, delta={delta}" , file=f)
    #print(f"(low) {low} <= {period} <= {high} (high)", file=f)
    mes = "    ==> "
    
        
    # Fraction が少ないときの補正 47:外部照射分割回数
    if int(col[47]) <= 10:
        high += 3
    if int(col[47]) <= 5:
        high += 3
           
    #if low <= period <= base + delta :
    print(f"low: {low} <= {period} <= {high} :high")
    if low <= period <= high :
        # OK
        print(f"{mes}frac:{frac}, days:{period} --- OK", file=f)
        s = f"{mes}frac:{frac}, days:{period} --- OK"
        err = 0
        # ここで、85:'放射線治療完遂度' のデータをチェックして、必要ならば、追記する。※特にSRC
        # 項目選択:['予定治療完遂',[予定治療完遂(8日以上の中断あり)',['予定の50%未満で中止',
        #          '予定の50%以上で中止','遂行程度不詳で中止','その他','不明']
        addText(col, low, period, high, '85:放射線治療完遂度', f) 
    else:
        # NG
        print(f"\n  #{n:04} dayCheck: *** ptr:{col[0]} *** ID:{id} 管理番号：{col[1]}  {col[3]} {col[2]}", file=f)
        print(f"st= {st}, en= {en}, period= {period}, frac= {frac}", file=f)
        print(f"(low) {low} <= {period} <= {high} (high) [orig:={base + delta}]", file=f)
        print(f"{mes}frac:{frac}, days:{period} *** NG ***", file=f)
        s = f"{mes}frac:{frac}, days:{period} *** NG ***"
        err = 1
        day_err += 1
        print(f'    === day_err: {day_err} ===', file=f)
        # ここで、85, '放射線治療完遂度' のデータをチェックして、必要ならば、追記する。
        addText(col, low, period, high, '85:放射線治療完遂度', f)
    return err, s, (low, period, high)


values = ['予定治療完遂','予定治療完遂(8日以上の中断あり)','予定の50%未満で中止',
          '予定の50%以上で中止','遂行程度不詳で中止','その他','不明']

# addText: low, period, high の内容に応じて、データを書き換え、その記録を残す
def addText(col, low, period, high, field, f):
    if field != '85:放射線治療完遂度':
        print(f'addText field error: {field}', file=f)
    # --- '85:放射線治療完遂度'
    ptr, ti = field.split(':')
    ptr = int(ptr)
    print(f'{field}: col: {col[ptr]}')

    if low <= period <= high:
        val = values[0] # normal
    elif period > high:
        val = values[1] # 予定治療完遂(8日以上の中断あり)
    else:
        val = 'Low days'
    val = f'{val} ** (<={col[ptr]})'
    print(f'val = [{val}]', file=f)
    return     


def pred(col, low, period, high, field, f=sys.stderr):
    """
    # 予想日数から、治療完遂を予測する。
    args
        col: 1 record data,
        low, period, high: 最小予想日数、実際の日数、最大予想日数
        field: 判定する項目 カラム番号（1始まり）：項目名
    return
        pre_comp: 予想完遂
    """
    if field != '85:放射線治療完遂度':
        print(f'addText field error: {field}', file=f)
    # --- '85:放射線治療完遂度'
    ptr, ti = field.split(':')
    ptr = int(ptr)
    #logger.debug(f'{field}: col: {col[ptr]}')

    if low <= period <= high:
        val = values[0] # normal
    elif period > high:
        val = values[1] # 予定治療完遂(8日以上の中断あり)
    else:
        val = 'Low days'
    val = f'**{val}** '
    #logger.debug(f'val = [{val}]', file=f)
    return  val   

# ---
if __name__ == "__main__":
    n = 2
    t = {1: '院内管理コード', 2: '性別', 3: '照射開始時年齢', 4: '照射開始時Karnofsky PS',
         5: '照射開始時ECOG PS', 6: '重複がん', 7: '重複がんの時期', 8: '重複がんメモ',
         9: '照射歴', 10: '疾患名', 11: '原発部位', 12: '原発部位側性',
         13: '原発部位ICD-Oコード', 14: '病理組織', 15: '病理組織ICD-Oコード',
         16: '病期分類名1', 17: 'CPR1', 18: 'T1', 19: 'N1', 20: 'M1', 21: 'Stage1',
         22: 'Grade1', 23: '病期分類名2', 24: 'CPR2', 25: 'T2', 26: 'N2', 27: 'M2',
         28: 'Stage2', 29: 'Grade2', 30: '病期分類名3', 31: 'CPR3', 32: 'T3', 33: 'N3',
         34: 'M3', 35: 'Stage3', 36: 'Grade3', 37: 'JASTRO 構造調査用疾患分類',
         38: '今回の治療', 39: '新患・再患', 40: '治療方針', 41: '併用療法', 42: '外来・入院',
         43: '外部照射開始日', 44: '外部照射終了日', 45: '外部照射総線量', 46: '外部照射日数',
         47: '外部照射分割回数', 48: '一日あたり照射回数', 49: '外部照射カテゴリー',
         50: '治療対象遠隔転移部位', 51: '外部照射部位', 52: '外部照射部位ICD-Oコード',
         53: '線種1', 54: 'エネルギー1', 55: '線種2', 56: 'エネルギー2', 57: '外部照射担当医',
         58: '外部照射指導医', 59: '特殊照射', 60: '治療加算1', 61: '放射線治療管理料一回目',
         62: '放射線治療管理料二回目', 63: '外部照射メモ', 64: '密封小線源部位', 65: '密封小線源部位ICD-Oコード',
         66: '密封線源', 67: '密封小線源線量率', 68: '密封小線源照射方法', 69: '密封小線源一回線量',
         70: '密封小線源分割回数', 71: '密封小線源総線量', 72: '密封小線源治療開始日',
         73: '密封小線源治療終了日', 74: '密封小線源治療日数', 75: '密封小線源担当医',
         76: '密封小線源指導医', 77: '三次元治療計画', 78: '密封小線源メモ', 79: '非密封線源',
         80: '非密封線源投与量', 81: '非密封線源投与回数', 82: '非密封線源投与日',
         83: '非密封線源担当医', 84: '非密封線源メモ', 85: '放射線治療完遂度', 86: '一次効果',
         87: '生死の状況', 88: '最終確認日', 89: '再発の有無', 90: '再発確認日', 91: '再発部位',
         92: '再発部位詳細', 93: '再発治療の有無', 94: '再発治療内容詳細', 95: '有害事象の有無',
         96: '有害事象確認日1', 97: '有害事象発生部位1', 98: '有害事象グレード1',
         99: '有害事象確認日2', 100: '有害事象発生部位2', 101: '有害事象グレード2',
         102: '有害事象確認日3', 103: '有害事象発生部位3', 104: '有害事象グレード3',
         105: '有害事象メモ', 106: '続発がんの有無', 107: '続発がん確認日', 108: '続発がん部位',
         109: '続発がんメモ', 110: '施設名', 111: '施設コード', 112: 'ID', 113: 'OBR_Set_ID',
         114: '診断タイプ', 'max_row': 21, 'max_column': 114}
    title = ['index']
    for i in range(1,115):
        title.append(t[i])
    col = [2, '16532', '女', '80', '', '', '', '', '', '', '肝癌', '肝', 'rt',
           'C22.0', 'Hepatocellular carcinoma; NOS', 'M8170/3',
           'UICC 8th Japanese', 'p', 'T1a', 'N0', 'M0', 'SIA', '', '', '', '',
           '', '', '', '', '', '', '', '', '', '', '', '乳癌', '新鮮', '新患',
           '術後', '', '外来', '2024/02/29', '2024/03/22', '42.56', '23', '16',
           '1', '', '', '乳頭部及び乳輪の悪性新生物', 'C50.0', 'Photon', '6',
           '', '', '', '', '', '画像誘導放射線治療加算', '複雑', '', '', '', '',
           '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
           '', '', '予定治療完遂', '', '非担癌生存', '2024/12/18', '', '', '',
           '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
           '', '足利赤十字病院', 'J09007', '0000041750', '', '原発巣']
    ok, s, (low, period, high)= dayCheckM(col, n, title, sys.stdout)
    print(f"\n#199: ok={ok}, mes='{s}', (low={low}, period={period}, high={high})")
    print("====================================================================")
    for p in [20, 24, 60]:
        print(f"period = {p}")
        addText(col, low, p, high, '85:放射線治療完遂度', sys.stdout)
        print("pred()=", pred(col, low, p, high, '85:放射線治療完遂度'))
    print("Normal end.")
