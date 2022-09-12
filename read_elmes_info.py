import sys, os, shutil, time, datetime
from file_handle import *


class ElmesInfo():
    def __init__(self, odr_data):
        self.odr_num = odr_data.odr_num
        self.root_dpath = r''
        self.set_ini_dpath(self.odr_num)
        self.set_ini_fname(self.odr_num)
        self.set_fpath(self.odr_num)

    def set_ini_dpath(self, odr_num):
        self.ini_dpath = odr_num

    def set_ini_fname(self, fname):
        self.ini_fname = fname
        self.ini_fpath = os.path.join(self.ini_dpath, self.ini_fname)

    def set_dpath(self, odr_num):
        self.dpath = os.path.join(self.root_dpath, f'{odr_num}')

    def get_master_data(self, kishu):
        pass

    def set_fpath(self, odr_num):
        self.fname = f'{odr_num}.txt'
        self.fpath = os.path.join(self.dpath, self.fname)

    def read_ini(self):
        with open(self.ini_fpath) as inif:
            inis = inif.readline()
        odr_ini, exm_ini = 0, 0
        for ini_head in inis:
            if ini_head == r'新設時工事番号':
                old_odr_row = ini_head.split('^t')[2]
            if ini_head == r'検討依頼番号':
                exm_row = ini_head.split('^t')[2]
            if (odr_ini == 1) and (exm_ini == 1):
                break
        return old_odr_row, exm_row

    def read_info(self):
        self.set_dpath('')
        self.set_fname()
        old_odr_row, exm_row = self.read_ini()
        with open(self.fpath) as f:
            lines = f.readlines()
            for l in lines:
                if l[0] == str(old_odr_row):
                    old_odr = l[1]
                if l[0] == str(exm_row):
                    exm_num = l[1]
        return old_odr, exm_num
