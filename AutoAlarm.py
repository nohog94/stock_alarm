from slacker import Slacker
import pymysql
import pandas as pd
import win32com.client
import ctypes
import time

################################################
# PLUS 공통 OBJECT
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


################################################
# PLUS 실행 기본 체크 함수
def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False

    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False

    # # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    # if (g_objCpTrade.TradeInit(0) != 0):
    #     print("주문 초기화 실패")
    #     return False

    return True


# 차트 기본 데이터 통신
class CpStockChart:
    def __init__(self):
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        self.objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

    def Request(self, code, cnt, objSeries):
        #######################################################
        # 1. 일간 차트 데이터 요청
        self.objStockChart.SetInputValue(0, code)  # 종목 코드 -
        self.objStockChart.SetInputValue(1, ord('2'))  # 개수로 조회
        self.objStockChart.SetInputValue(4, cnt+1)  # 최근 cnt 일치
        self.objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 날짜,시가,고가,저가,종가,거래량
        self.objStockChart.SetInputValue(6, ord('D'))  # '차트 주기 - 일간 차트 요청
        self.objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.objStockChart.BlockRequest()

        rqStatus = self.objStockChart.GetDibStatus()
        rqRet = self.objStockChart.GetDibMsg1()
        #print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()
        time.sleep(0.5)
        #######################################################
        # 2. 일간 차트 데이터 ==> CpIndexes.CpSeries 로 변환
        len = self.objStockChart.GetHeaderValue(3)

        # for i in range(len):
        #     day = self.objStockChart.GetDataValue(0, len - i - 1)
        #     open = self.objStockChart.GetDataValue(1, len - i - 1)
        #     high = self.objStockChart.GetDataValue(2, len - i - 1)
        #     low = self.objStockChart.GetDataValue(3, len - i - 1)
        #     close = self.objStockChart.GetDataValue(4, len - i - 1)
        #     vol = self.objStockChart.GetDataValue(5, len - i - 1)
        #      print(day, open, high, low, close, vol)
        #     # objSeries.Add 종가, 시가, 고가, 저가, 거래량, 코멘트
        #     objSeries.Add(close, open, high, low, vol)

        close_total = [None] * len
        for i in range(len):
            close = self.objStockChart.GetDataValue(4, i)
            close_total[i] = close

        m26_close = 0
        m26_close_prev = 0
        m12_close = 0
        m12_close_prev = 0
        for i in range(0, 26):
            m26_close += close_total[i]
        for i in range(1, 27):
            m26_close_prev += close_total[i]
        for i in range(0, 12):
            m12_close += close_total[i]
        for i in range(1, 13):
            m12_close_prev += close_total[i]

        m26_close /= 26
        m26_close_prev /= 26
        m12_close /= 12
        m12_close_prev /= 12
        if m26_close < m12_close and m26_close_prev > m12_close_prev:
            return True

        return False


class AutoAlarm:
    def __init__(self):
        self.conn = pymysql.connect(host='localhost', user='root',
                                    password='11hello', db='INVESTAR', charset='utf8')

        sql = "SELECT code, company FROM company_info"
        self.codes = pd.read_sql(sql, self.conn)

    def __del__(self):
        """소멸자: MariaDB 연결 해제"""
        self.conn.close()

    def send_alarm(self, company):
        slack = Slacker('xoxb-1651075977093-1651091890133-WMgUXyqKu8uywGWo2uhTdNYe')
        # Send a message to #general channel
        slack.chat.post_message('#test', f'{company} is on signal!')

if __name__ == '__main__':
    alarm = AutoAlarm()
    objChart = CpStockChart()
    objSeries = win32com.client.Dispatch("CpIndexes.CpSeries")
    for i in range(alarm.codes.shape[0]):
        code = alarm.codes.iloc[i,0]
        company = alarm.codes.iloc[i,1]
        rcode = 'A' + code
        try:
            signal = objChart.Request(rcode, 26, objSeries)
            if signal:
                alarm.send_alarm(company)
        except:
            print(code)
    # objMarketTotal = CMarketTotal()
    # objMarketTotal.GetAllMarketTotal()


