import sys, os, shutil, glob, datetime
import read_config

global odr_list
global dwg_list
global doc_list
global root_dpath
global exe_dpath
global tmp_dl_dpath
global result_dpath
global tool_fpath
global result_name_format
global ver
global debug


class SetDefaultVar:
    def __init__(self, conf_f):
        self.conf_f = conf_f

    def parse_sys(self, args):
        global odr_list
        global dwg_list
        global doc_list
        global exe_dpath
        self.args = args
        if len(args) > 2:
            exe_dpath = os.getcwd()
            odr_list = self.args[1]
            dwg_list = self.args[2]
            doc_list = self.args[3]
            if debug == 0:
                exe_dpath = os.path.dirname(self.args[0])
                odr_list = odr_list.strip('[').strip(']').split(',')[1:]
                dwg_list = dwg_list.strip('[').strip(']').split(',')
                doc_list = list(map(int, doc_list.strip('[').strip(']').split(',')))
                print(exe_dpath)
                print(odr_list)
                print(dwg_list)
                print(doc_list)
        else:
            sys.exit()

    def set_variables(self):
        global root_dpath
        # global exe_dpath
        global tmp_dl_dpath
        global result_dpath
        global tool_fpath
        global result_name_format
        global ver

        # 設定ファイルの読み込み
        conf = read_config.ReadConfig(self.conf_f)
        conf.set_conf_category('DEBUG')
        root_dpath = str(conf.set_var('ROOT_DIR'))
        # exe_dpath = str(conf.set_var('EXE_DIR'))
        result_dname = str(conf.set_var('RESULT_DIR'))
        tmp_dl_dname = str(conf.set_var('TMP_DIR'))
        result_name_format = str(conf.set_var('RESULT_NAME'))
        ver = str(conf.set_var('Ver'))

        # ファイル名設定
        # # UIエクセル
        tool_fname = f'工事ファイル収集ツール_Ver{ver}.xlsm'
        # フォルダ設定
        tmp_dl_dpath = os.path.join(exe_dpath, tmp_dl_dname)
        result_dpath = os.path.join(exe_dpath, result_dname)

        # フルパス設定
        # # ツール
        tool_fpath = os.path.join(exe_dpath, tool_fname)

class SystemInitialize:
    def __init__(self, root_dirs=[], if_dirs=[]):
        self.root_dirs = root_dirs
        self.if_dirs = if_dirs
        self.check_root_dir()
        self.mk_if_dirs()
        self.rm_if_files()

    def rm_files_in_dir(self, dpath, fname='*'):
        rm_files = glob.glob(os.path.join(dpath, fname))
        for f in rm_files:
            if os.path.isfile(f): os.remove(f)

    def rm_if_files(self):
        for d in self.if_dirs:
            self.rm_files_in_dir(d)

    def mk_if_dirs(self):
        for d in self.if_dirs:
            if not os.path.exists(d):
                os.mkdir(d)

    def check_root_dir(self):
        for d in self.root_dirs:
            if not os.path.exists(d):
                sys.exit()


debug = 0
conf_f = 'collect_docs.ini'
if debug ==0:
    args = sys.argv # 引数：[本体実行パス, 検索サイト, 工事番号リスト, 図番リスト]
elif debug == 1:
    exe_dpath = r'.'
    odr_list = ['POY36218','PYWQ4214', 'PX012501']
    # odr_list = ['MKD269061', 'MZRG1405', 'MBE31410', 'MZRG5505']
    dwg_list = ['昇降路平面図', 'その他(エレベーター)']
    doc_list = list(map(int, ['3']))
    args = ['path', odr_list, dwg_list, doc_list]
set_vars = SetDefaultVar(conf_f)
set_vars.parse_sys(args)
set_vars.set_variables()
ini_f = SystemInitialize([root_dpath, exe_dpath], [tmp_dl_dpath, result_dpath])
