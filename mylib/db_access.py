#### !python -m pip install pyodbc
"""
Connects to a SQL database using pyodbc
"""
import pyodbc
import pprint

#SERVER = '<server-address>'
#DATABASE = '<database-name>'
#USERNAME = '<username>'
#PASSWORD = '<password>'

SERVER = '172.31.12.11'
DATABASE = 'MOSAIQ'
USERNAME = 'MOSAIQUser'
PASSWORD = 'mosaiq'

connectionString = None

def db_init():
    global connectionString
    
    driver = "{SQL Server}"
    connectionString = f'DRIVER={driver};' +f'SERVER={SERVER};' +\
                    f'DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};Trusted_Connection=no'
    return connectionString
    #return


# SQL 実行、結果を印字
def query(sql):
    conn = pyodbc.connect(connectionString)
    cursor = conn.cursor()

    cursor.execute(sql) # SQL実行
    
    rows = cursor.fetchall()
    print("len-rows=", len(rows))
    cursor.close()
    conn.close()
    return rows

# transaction
def DBtrans(sql, values):
    print("### DBtrans ###")
    print("sql=", sql)
    print("values=", values)
    
    try:
        # 接続
        conn = pyodbc.connect(connectionString)
        cursor = conn.cursor()
    
        # トランザクション開始
        #cursor.execute("BEGIN TRANSACTION")
    
        # SQL文の実行
        cursor.execute(sql, values)
    
        # コミット
        conn.commit()
        print("データが更新されました")
    
    except pyodbc.Error as ex:
        print('Error:', ex)
        # ロールバック
        conn.rollback()
    
    finally:
        # 接続を閉じる
        if conn:
            conn.close()
    return
#---
# ID=16421 の病態、最終生存確認日 UPDATE
def execSQL(sql, values):
    conn = pyodbc.connect(connectionString)
    cursor = conn.cursor()

    sql = '''update admin set user_defined_pro_id_3 = ?
    where pat_id1 = ? ;'''
    
    values=(13114, 16839) # 現病死、 ID=16421
    print("SQL=", sql)
    print("values=", values)

    try:
        cursor.execute(sql, values)
        conn.commit()
    except Exception as e:
        print(f'error: {e}')
        conn.rollback()
    finally:
        cursor.close()
        conn.close()
    print("実行しました.")
    
    
    ### 出力
    '''
    SQL= update admin set user_defined_pro_id_3 = ?
    where pat_id1 = ? ;
    values= (13114, 16839)
    実行しました.
    '''

if __name__ == "__main__":
    db_init()
    SQL_QUERY = '''select pat_id1,user_defined_dttm_1,user_defined_pro_id_3 
    from admin where pat_id1 = '16119';'''
    rows = query(SQL_QUERY)
    pprint.pprint(rows)

    # [(16119, datetime.datetime(2023, 4, 29, 0, 0), 13114)]
    # 管理番号、最終生存確認日、病態


    status = {'13111':'1.非担癌生存','13114':'4.原病死', '13113':'3.担癌不詳生存', '13112':'2.担癌生存','13115':'5.他病死',
          '13116':'6.不明死', '13117':'7.消息不明' }
    st=rows[0][2]
    print(st)
    print(status[str(st)])  
    # 13114
    # 4.原病死


    # ID= 16119 の カルテ番号、氏名を表示
    sql = '''
    select i.ida,p.last_name,p.first_name from ident i 
    left outer join patient p on i.pat_id1=p.pat_id1 
    where i.pat_id1=16839 ;'''
    print("SQL=", sql)
    query(sql)

    ### ---- 出力例
    '''
    SQL= 
    select i.ida,p.last_name,p.first_name from ident i 
    left outer join patient p on i.pat_id1=p.pat_id1 
    where i.pat_id1=16839 ;
    [('', '足利', 'update')]
    '''

    # 性別、生年月日 を表示
    sql= """select i.Pat_Id1,i.ida,p.Last_Name,p.First_Name,a.Gender,
                dbo.fn_GetPatientAge(i.Pat_Id1,GETDATE()) from ident i
                left outer join patient p on i.Pat_Id1=p.Pat_ID1
                left outer join admin a on a.Pat_ID1=i.Pat_Id1
                where i.ida='00000202308' ;"""

    print("SQL=", sql)
    query(sql)




    # ID=16421 の病態、最終生存確認日 を表示  --- OK
    sql = '''select pat_id1,user_defined_dttm_1,user_defined_pro_id_3 
    from admin where pat_id1 = '16839';'''

    print("SQL=", sql)
    query(sql)# ID=16421 の病態、最終生存確認日 UPDATE
    #### 出力
    '''SQL= select pat_id1,user_defined_dttm_1,user_defined_pro_id_3 
    from admin where pat_id1 = '16839';
    [(16839, None, 13114)]'''








    #############  TEST ######################

    sql='''update admin set user_defined_pro_id_3 = 13114
    where pat_id1 = 16839 ;'''

    #values=(13113, 16839) 
    DBtrans(sql)

    ### 出力
    '''### trans2 ###
    sql= update admin set user_defined_pro_id_3 = 13114
    where pat_id1 = 16839 ;
    values= (13113, 16839)
    データが更新されました'''


