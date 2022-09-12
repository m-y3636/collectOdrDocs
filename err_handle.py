import sys, os, shutil, re
import logging
import read_config as rconf
from functools import wraps
from set_variables import conf_f


err_file = r'err.txt'
err_conf = rconf.ReadConfig(conf_f)
err_conf.set_conf_category('ErrCode')


class ExpLog():
    def __init__(self, log_path, log_format):
        self.log_path = log_path
        self.log_format = log_format

        self.log = logging.getLogger(__name__)
        log_formatter = logging.Formatter(
            log_format
        )
        # '%(asctime)s <%(levelname)s> : []%(message)s'
        filehandler = logging.FileHandler(log_path, mode='a')
        filehandler.setFormatter(log_formatter)
        self.log.setLevel(logging.DEBUG)
        self.log.addHandler(filehandler)

    def logf_exp(self, *args, **kwargs):
        message = ','.join(args)
        self.log.debug(message)
        # self.log.debug(args[0], args[1:])


class ErrHandle:
    def err_handle(err_val):
        def _except_dec(func):
            @wraps(func)
            def _except_internal(*args, **kwargs):
                try:
                    return func(*args, **kwargs)
                except NameError:
                    err_type = '変数定義エラー'
                except TypeError:
                    err_type = 'データ型エラー'
                except Exception as e:
                    args[0].prog_odr.result_dwgsender = False
                    err_type = str(e)
                err_process = err_conf.basic_conf.get(str(err_val))
                logger.logf_exp(f'[{str(err_val)}]{err_process}:',
                                args[0].site_name,
                                args[0].prog_odr.odr_num,
                                args[0].prog_odr.odr_name,
                                args[0].prog_odr.exm_num,
                                args[0].prog_odr.dwg_num,
                                args[0].prog_odr.dwg_name,
                                err_type)
                # logger.logf_exp(f'[{str(err_val)}]{err_process}:', err_type)
                # logger.logf_exp(f'[{str(err_val)}]{err_conf.basic_conf.get(str(err_val))}:', err_type)
                return err_val
            return _except_internal
        return _except_dec


@ErrHandle.err_handle(err_val=200)
def print_test(check):
    print_internal()
    return 3

def print_internal():
    print(test)
    return 1

logger = ExpLog(err_file, '%(asctime)s [%(levelname)s] %(message)s')


if __name__ == '__main__':
    print(print_test('Hello'))