import sys, os, shutil, glob, datetime, time, re
from set_variables import tmp_dl_dpath, result_dpath


class MyFile:
    def __init__(self):
        self.fpath = ''
        self.f_name = ''
        self.dpath = ''

    def set_fpath(self, fpath):
        self.fpath = fpath
        self.fname = os.path.basename(fpath)
        self.dpath = os.path.dirname(fpath)

    def set_dpath(self, dpath):
        self.dpath = dpath
        if not self.fname == '': self.fpath = os.path.join(dpath, self.fname)

    def set_fname(self, fname):
        self.fname = fname
        if not self.dpath == '': self.fpath = os.path.join(self.dpath, self.fname)

    def open_data(self):
        with open(self.path) as f:
            self.lines = f.readline()

    def set_fname_datetime(self, datetime):
        self.f_name_dt = os.path.splitext(self.f_name)[0] + "_" + datetime.strftime("%Y%m%d") + "_" + os.path.splitext(self.f_name)[1]
        self.path = os.path.join(self.dir_name, self.f_name_dt)

    def copy_to(self, cppath):
        shutil.copy2(self.fpath, cppath)


class OdrFile(MyFile):
    def set_dwg_data(self, dwg_num, dwg_rev, dwg_name, odr_num, odr_name, doc_type):
        if (self.site_name == '図面伝送') or (self.site_name == 'ELENA-Ⅱ'):
            if dwg_rev == r'*': dwg_rev = r'@'
        self.set_fname(f'{dwg_num}_{dwg_rev}.{doc_type}')
        self.set_dpath(tmp_dl_dpath)
        self.result_fname = f'{dwg_num}_{dwg_rev}_{dwg_name}_{self.site_name}_{odr_num}.{doc_type}'
        self.result_fname = re.sub(r'[\\|/|:|?|"|<|>|\|]', '', self.result_fname)
        self.result_dpath = os.path.join(result_dpath, re.sub(r'[\\|/|:|?|"|<|>|\|]', '', f'{odr_name[:30]}_{odr_num}'))
        self.result_fpath = os.path.join(self.result_dpath, self.result_fname)

    def mv_to_resultd(self):
        if not os.path.exists(self.result_dpath): os.mkdir(self.result_dpath)
        for i in range(20):
            if os.path.exists(self.fpath):
                shutil.copy2(self.fpath, self.result_fpath)
                break
            time.sleep(1)
        if os.path.exists(self.result_fpath):
            os.remove(self.fpath)
            self.set_fpath(self.result_fpath)

    def set_dl_site(self, site_name):
        self.site_name = site_name


