from slacker import Slacker
import pymysql
import pandas as pd
import win32com.client
import ctypes
import time


# 차트 기본 데이터 통신
class CpStockChart:
    def __init__(self, day):
        self.day = day
        self.conn = pymysql.connect(host='localhost', user='root',
                                    password='11hello', db='INVESTAR', charset='utf8')

        sql = f"SELECT code, date, open, close FROM daily_price where date > date_add(SYSDATE(), INTERVAL -{day} DAY) "
        self.prices = pd.read_sql(sql, self.conn)

        sql = "SELECT * FROM company_info"
        self.company = pd.read_sql(sql, self.conn)

    def __del__(self):
        """소멸자: MariaDB 연결 해제"""
        self.conn.close()

    def back_test(self):
        earn = 0
        tax = 0.2
        commission = 0.015
        cnt = 0
        for code in self.company['code']:
            a = self.prices[self.prices['code'] == code]
            a.reset_index(inplace=True)
            mav26 = a['close'].rolling(window=26).mean()
            mav12 = a['close'].rolling(window=12).mean()
            for i in range(len(mav12) - 1):
                if mav12.iloc[i] < mav26.iloc[i] and mav12.iloc[i + 1] > mav26.iloc[i + 1]:
                    if i < len(mav12) - 2:
                        if a['open'][i + 2] == 0:
                            continue
                        get = (a['close'][i + 2] - a['open'][i + 2]) / a['open'][i + 2] * 100
                        loss = commission + (100+get) * (tax/100) + (100+get) * (commission/100)
                        earn = earn + get - loss
                        cnt += 1
                        print(code + '\t' + str(a['open'][i + 2]) + '\t' + str(a['close'][i + 2]) + '\t' + str(
                            get) + '\t' + str(earn) + '\t' + str(cnt))
                    else:
                        continue

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
    day = 200
    objChart = CpStockChart(day)
    objChart.back_test()



