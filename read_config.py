import configparser as confp

global root_dir
global new_dl_dname
global template_dname
global last_dname
global result_dname

class ReadConfig:
    def __init__(self, conf_f):
        # 設定ファイルの読み込み：DLフォルダ、UIファイル、URL設定
        self.conf = confp.ConfigParser()
        self.conf.read(conf_f, encoding='utf-8')

    def set_conf_category(self, category='DEBUG'):
        self.basic_conf = self.conf[category]

    def set_var(self, key):
        return str(self.basic_conf.get(key))


# conf_f =
# class ReadCollectDocsIni(ReadConfig):
#     def __init__(self):