# -*- coding: utf-8 -*-
import sys
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QAxContainer import *
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QPushButton
from pandas import Series, DataFrame
import pandas as pd
import numpy as np
import threading
import json
import ctypes, sys

#import flask
import cherrypy

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


class IIndi(QAxWidget):
    # OCX_ADDR = 'GIEXPERTCONTROL64.GiExpertControl64Ctrl.1'
    OCX_ADDR = 'GIEXPERTCONTROL.GiExpertControlCtrl.1'
    tr_map = {}
    fields_map = {'stock_mst': ['STD_CODE', 'SHORT_CODE', 'MARKET_TYPE', 'STOCK_NAME', 'SECTOR', 'STMT_DATE', 'SUSP_TYPE', 'MGMT_TYPE', 'ALERT', 'DECL_TYPE', 'MARGIN_TYPE', 'CREDIT_TYPE', 'ETF_TYPE', 'SEC_TYPE']}
    def __init__(self):
        super().__init__(self.OCX_ADDR)
        self.req_map = {}
        self.res_map = {}
        self.ReceiveData.connect(self.recv_data)
        self.ReceiveSysMsg.connect(self.recv_msg)
    def set_query_name(self, query):
        return self.dynamicCall('SetQueryName(QString)', query)
    def request_data(self):
        return self.dynamicCall('RequestData()')
    def get_multi_row_count(self):
        return self.dynamicCall('GetMultiRowCount()')
    def set_single_data(self, index, string):
        return self.dynamicCall('SetSingleData(int, QString)', index, string)

    def req_stock_mst(self):
        self.set_query_name('stock_mst')
        req_id = self.request_data()
        self.req_map[req_id] = 'stock_mst'
        with cv:
            cv.wait()
        return req_id

    def recv_stock_mst(self):
        row_cnt = self.GetMultiRowCount()
        rows = []
        for i in range(row_cnt):
            row = {}
            for j in range(len(self.fields_map['stock_mst'])):
                row[self.fields_map['stock_mst'][j]] = self.GetMultiData(i, j)
            rows.append(row)
        return json.dumps(rows)

    def recv_data(self, req_id):
        tr_name = self.req_map[req_id]
        res = self.tr_map[tr_name](self)
        del self.req_map[req_id]
        self.res_map[req_id] = res
        with cv:
            cv.notify()
 
    def recv_msg(self, id):
        print('sysmsg id', id)
    tr_map['stock_mst'] = recv_stock_mst
 

class IIndiWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.iindi = IIndi()
        self.setWindowTitle("IIndi")

#app = flask.Flask(__name__)
#@app.route('/')
class Router:
    @cherrypy.expose
    def index(self):
        req_id = iindi_window.iindi.req_stock_mst()
        res = iindi_window.iindi.res_map[req_id]
        del iindi_window.iindi.res_map[req_id]
        return res

cv = threading.Condition()

if __name__ == "__main__":
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit()

    qt_app = QApplication(sys.argv)
    iindi_window = IIndiWindow()
    iindi_window.show()
    threading.Thread(target=cherrypy.quickstart, args=(Router(),), daemon=True).start()
    qt_app.exec()
