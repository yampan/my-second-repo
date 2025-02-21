# =====================================================================
# Original: Pi5a:/home/jupyter/work/Oauth/common_lib/logger.py
# 2025-02-04
# 使用方法：
#     sys.path.append('/home/jovyan/work/OAuth/common_lib')
#     from logger import FMT, FMT2, createLogger, clearLogfile, log_init
#
#     初期化
#       os.makedirs("./log", exist_ok=True)
#       LOG_FN = "LOG_JROD.TXT"
#       logger = init_log(LOG_FN)
#
# =====================================================================
# Logger用： 時刻をJSTにする。
### example format  2024-10-28
from datetime import datetime, timedelta, timezone
from logging import Formatter, LogRecord

class DatetimeFormatter(Formatter):
    def formatTime(self, record: LogRecord, datefmt=None):
        if datefmt is None:
            #datefmt = "%Y-%m-%d %H:%M:%S.%03d" # logging.Formatterのデフォルトと同じ形式
            datefmt = "%Y-%m-%d %H:%M:%S.%f" # for windows
            #datefmt = "%Y-%m-%d %H:%M:%S.%f%z" # 2024-08-15 18:09:17.468103+0900
        TZ_JST = timezone(timedelta(hours=+9), 'JST')
        created_time = datetime.fromtimestamp(record.created, tz=TZ_JST)
        s = created_time.strftime(datefmt)
        return s[:23] # 23桁まで

#fmt = DatetimeFormatter("%(asctime)s %(name)-8s:%(lineno)-3s %(funcName)s [%(levelname)-7s]: %(message)s") # ここでフォーマットを指定する
FMT = DatetimeFormatter("%(asctime)s [%(levelname)-7s] %(name)s|%(funcName)s:%(lineno)-3s  %(message)s") # ここでフォーマットを指定する
FMT2 = DatetimeFormatter("%(asctime)s [%(name)s] %(message)s") 

# sh.setFormatter(fmt) で設定する。

#format = "%(levelname)-9s  %(asctime)s [%(filename)s:%(lineno)d] %(message)s"
format = "%(asctime)s - %(name)s - %(message)s"
# ハンドラ自体にフォーマットを設定しておく必要がある。
# format = "%(levelname)-9s  %(asctime)s [%(filename)s:%(lineno)d] %(message)s"
# fl_handler.setFormatter(Formatter(format))

# ------------------
### Logger  2024-10-28
from logging import getLogger, StreamHandler, Formatter, DEBUG, WARNING, handlers
#import pickleshare
import shutil, os

def createLogger(id='log', LOG_FN="LOG_C.TXT", rotate=False,
                                     propagate=True, format=None, deb=0):
    # create logger
    if deb: print(f"createLogger: id={id}, LOG_FN={LOG_FN}, rotate={rotate},", 
                                      f"propagate={propagate}")
    logger = getLogger(id)
    logger.setLevel(DEBUG)
    
    # create console handler and set level to debug
    ch = StreamHandler()
    ch.setLevel(WARNING)
    fmt0 = DatetimeFormatter("%(asctime)s - %(name)s - %(message)s")
    if deb: print("type-fmt0=", type(fmt0))
    ch.setFormatter(fmt0)
    
    # create formatter
    formatter = Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')   
    fh = handlers.RotatingFileHandler(LOG_FN, encoding="utf-8",
                                     backupCount=5,maxBytes=1024*100) # 100 KB   
    if format is not None:
        if deb: print("   createLogger: fmt: format")
        fh.setFormatter(format)
    else:
        if deb: print("   createLogger: fmt: defaut")
        fh.setFormatter(formatter)
    fh.setLevel(DEBUG)
    if rotate:
        fh.doRollover()
    # add handler to logger
    logger.addHandler(ch)
    logger.addHandler(fh)
    logger.propagate = propagate
    return logger
    
# --- ファイルを初期化
def clearLogfile(LOG_FN):
    global REC_NO
    print(f"   *** LOG FILE({LOG_FN}) CLEAR ***")
    #log_fn = "LOG_C.TXT"
    log_fn = LOG_FN
    if os.path.isfile(log_fn):  os.remove(log_fn)
    shutil.rmtree("./log")
    os.makedirs("./log", exist_ok=True)
    REC_NO = 0
    
# ============ Logger start =============
def log_init(LOG_FN):
    #LOG_FN = 'LOG_C.TXT'
    clearLogfile(LOG_FN)
    logger = createLogger("log", LOG_FN, format=FMT)
    
    logger.debug('debug message')
    logger.info('info message')
    logger.warning('warn message')
    logger.error('error message')
    logger.critical('critical message')
    print("root-log created. END.")
    return logger

import glob
def get_file_info(directory, pattern, show=0):
    """
    指定したディレクトリ内で、パターンに一致するファイルの情報を取得する。

    Args:
        directory (str): ディレクトリのパス。
        pattern (str): ファイル名のパターン。

    Returns:
        list: ファイル情報のリスト。
    """

    file_info_list = []
    file_paths = glob.glob(os.path.join(directory, pattern))  # パターンに一致するファイルパスを取得

    for file_path in file_paths:
        file_name = os.path.basename(file_path) # ファイル名のみ抽出
        file_size = os.path.getsize(file_path)
        file_ctime = os.path.getctime(file_path)  # 作成日時 (Windows)

        file_ctime_datetime = datetime.fromtimestamp(file_ctime)

        file_info = {
            "name": file_name,
            "size": file_size,
            "ctime": file_ctime_datetime.strftime("%Y-%m-%d %H:%M:%S"),
        }
        file_info_list.append(file_info)

    # if show:
    print("-" * 60)  
    for file_info in file_info_list:
        print(f"{file_info['name']:30} {file_info['size']:6} {file_info['ctime']}")
    print("-" * 60)    

    return file_info_list


# ---- for debug
if __name__ == "__main__":
    print("このスクリプトは直接実行されました")
    # logger start
    os.makedirs("./log", exist_ok=True)
    LOG_FN = "LOG_JROD.TXT"
    clearLogfile(LOG_FN)
    logger = createLogger("log", LOG_FN, format=FMT)
    logger.debug('debug message')
    logger.info('info message')
    logger.warning('warn message')
    logger.error('error message')
    logger.critical('critical message')    

    dir = "."
    pattern = "LOG*"
    fns = get_file_info(dir, pattern, show=1)
    # ファイル情報を表示
    for file_info in fns:
        print(f"Name: {file_info['name']}")
        print(f"Size: {file_info['size']} bytes")
        print(f"Created Time: {file_info['ctime']}")
        print("-" * 20)    