from WindPy import *
import numpy as np
import pandas as pd
import datetime
import math
import xlswr
import xml.dom.minidom
import os

class Strategy:
    global w
    #买信号（dict）
    buy_signal = {}
    #卖信号（dict）
    sell_signal = {}
    #关注的年报年份
    year = 0
    #开始时间
    start_date = ""
    #结束时间
    end_date = ""
    #交易日历
    trade_calendar = []
    #买入信号最后生成日期
    last_signal_date = ""
    #清盘信号生成日期
    clear_signal_date = ""
    #持仓（dict)
    position = {}
    #初始资产净值
    initial_asset_value = 0
    #现金
    cash = 0
    #净值表（list)
    asset_value = []
    #手续费
    commission = 0
    #交易记录(list)
    transaction = []

    def initialize(self):
        #策略参数初始化
        dom = xml.dom.minidom.parse("strategyConfig.xml")
        root = dom.documentElement
        self.year = int(root.getElementsByTagName("year")[0].firstChild.nodeValue)
        self.last_signal_date = root.getElementsByTagName("last_signal_date")[0].firstChild.nodeValue
        self.clear_signal_date = root.getElementsByTagName("clear_signal_date")[0].firstChild.nodeValue
        self.start_date = root.getElementsByTagName("start_date")[0].firstChild.nodeValue
        self.end_date = root.getElementsByTagName("end_date")[0].firstChild.nodeValue
        self.initial_asset_value = float(root.getElementsByTagName("initial_asset_value")[0].firstChild.nodeValue)
        self.cash = self.initial_asset_value
        self.commission = float(root.getElementsByTagName("commission")[0].firstChild.nodeValue)

        #交易日历初始化
        try:
            self.trade_calendar = xlswr.readXls2List("交易日历.xls", "交易日")
            self.trade_calendar = [str(item) for item in self.trade_calendar]
        except:
            trade_days = w.tdays(self.start_date, self.end_date, "").Data[0]
            for trade_day in trade_days:
                self.trade_calendar.append(datetime.datetime.strftime(trade_day, '%Y%m%d'))
            xlswr.writeList2Xls(self.trade_calendar, ["交易日"], "交易日历.xls")

    # 加载策略持仓
    def loadPosition(self, date):
        if os.path.exists("./" + date + "/持仓.xls") == False:
            if date < self.trade_calendar[0]:
                self.position = dict()
            else:
                print("未检测到" + date + "的日终持仓文件\n")
                exit(1)

        else:
            try:
                self.position = xlswr.readXls2Dict("./" + date + "/持仓.xls", 0)
            except:
                print("读取" + date + "的日终持仓文件失败\n")
                exit(1)

    # 加载买入信号
    def loadBuySignal(self, date):
        if os.path.exists("./" + date + "/买入信号.xls") == False:
            print(date + "日终并未生成新的买入信号\n")
            self.buy_signal = dict()
        else:
            try:
                self.buy_signal = xlswr.readXls2Dict("./" + date + "/买入信号.xls", 0)
            except:
                print("读取" + date + "日终生成的买入信号文件失败\n")
                exit(1)

    # 加载卖出信号
    def loadSellSignal(self, date):
        if os.path.exists("./" + date + "/卖出信号.xls") == False:
            print(date + "日终并未生成新的卖出信号\n")
            self.sell_signal = dict()
        else:
            try:
                self.sell_signal = xlswr.readXls2Dict("./" + date + "/卖出信号.xls", 0)
            except:
                print("读取" + date + "日终生成的卖出信号文件失败\n")
                exit(1)

    # 加载净值文件，获取可用现金
    def loadAssetValue(self, date):
        file_path = "./" + date + "/净值.xls"
        if os.path.exists(file_path) == False:
            if date < self.trade_calendar[0]:
                self.asset_value = []
            else:
                print("未检测到" + date + "的日终净值文件\n")
                exit(1)
        else:
            try:
                data = pd.read_excel(file_path)
                data["日期"] = data["日期"].astype(np.str)
                data["单位净值"] = data["单位净值"].astype(np.float)
                data["资产净值"] = data["资产净值"].astype(np.float)
                data["可用现金"] = data["可用现金"].astype(np.float)
                self.asset_value = data.values.tolist()
                self.cash = self.asset_value[-1][-1]
            except:
                print("读取截止到" + date + "日终的净值文件失败\n")
                exit(1)


    def dailyProcess(self):
        today = datetime.datetime.now()
        #today = datetime.datetime(2017, 1, 3)
        #today = date
        yesterday = today + datetime.timedelta(days = -1)
        today = today.strftime("%Y%m%d")
        yesterday = yesterday.strftime("%Y%m%d")
        prev_trade_date = w.tdaysoffset(-1, today, "").Data[0][0].strftime("%Y%m%d")

        if os.path.exists(today) == False:
            os.mkdir(today)

        #若当前日期早于1月第一个交易日，生成买入信号
        if today <= self.last_signal_date:
            self.generateBuySignal(today)
            return

        # 加载前一交易日日终持仓
        self.loadPosition(prev_trade_date)

        # 加载昨日日终生成的买入信号
        self.loadBuySignal(yesterday)

        # 加载昨日日终生成的卖出信号
        self.loadSellSignal(yesterday)

        # 加载截至前一交易日日终的净值文件，获得可用现金
        self.loadAssetValue(prev_trade_date)

        # 若当日为交易日，执行前一日日终生成的买卖信号
        if today in self.trade_calendar:
            #执行前一日日终生成的买卖信号
            self.order(today)
            if today == self.trade_calendar[0]:
                self.initial_asset_value -= self.cash
                self.cash = 0
                dom = xml.dom.minidom.parse("strategyConfig.xml")
                root = dom.documentElement
                node = root.getElementsByTagName("initial_asset_value")[0].firstChild
                node.nodeValue = self.initial_asset_value
                with open("strategyConfig.xml", 'wb+') as f:
                    f.write(dom.toprettyxml(indent='\t', encoding='utf-8'))
            #日终估值
            self.asset_evaluation(today)

        # 若当前日期为3月倒数第二个交易日，生成清盘卖出信号；否则，正常生成卖出信号
        if today == self.clear_signal_date:
            self.generateClearSignal(today)
        else:
            self.generateSellSignal(today)
        print("完成" + today + "的全部处理！\n")

    def order(self, date):
        if not self.buy_signal and not self.sell_signal:
            return

        stock_codes = list(self.buy_signal.keys()) + list(self.sell_signal.keys())
        trade_status = w.wss(stock_codes, "trade_status", "tradeDate=" + date).Data[0]
        maxupordown = w.wss(stock_codes, "maxupordown", "tradeDate=" + date).Data[0]
        trade_status = pd.Series(trade_status, index = stock_codes)
        maxupordown = pd.Series(maxupordown, index = stock_codes)
        open_prices = w.wss(stock_codes, "open", "tradeDate=" + date + ";priceAdj=U;cycle=D").Data[0]
        open_prices = pd.Series(open_prices, index = stock_codes)
        open_prices_f = w.wss(stock_codes, "open", "tradeDate=" + date + ";priceAdj=F;cycle=D").Data[0]
        open_prices_f = pd.Series(open_prices_f, index = stock_codes)
        other_prices = w.wss(stock_codes, "close, high, low", "tradeDate=" + date + ";priceAdj=U;cycle=D").Data
        other_prices = pd.DataFrame(data = np.matrix(other_prices).T, index = stock_codes, columns = ["close", "high", "low"])

        #处理卖信号
        for stock_code in list(self.sell_signal.keys()):
            open_price = open_prices[stock_code]
            high_price = other_prices.at[stock_code, "high"]
            low_price = other_prices.at[stock_code, "low"]
            if trade_status[stock_code] == '交易' and (maxupordown[stock_code] == 0 or (maxupordown[stock_code] != 0 \
            and (open_price != low_price or open_price != high_price))):
                s = self.sell_signal[stock_code]
                stock_name = s[0]
                amount = s[1]
                self.cash = self.cash + open_price * amount * (1 - self.commission)
                del self.sell_signal[stock_code]
                del self.position[stock_code]
                # 记录交易，包括日期、证券代码、交易市场代码、交易方向、交易数量、交易价格
                tmp = stock_code.split('.')
                trade_code = tmp[0]
                market = tmp[1]
                if market == 'SZ':
                    market = 'XSHE'
                else:
                    market = 'XSHG'
                self.transaction.append([date, "09:30:00", trade_code, market, "SELL", '', amount, open_price])

        #处理买信号
        for stock_code in list(self.buy_signal.keys()):
            open_price = open_prices[stock_code]
            high_price = other_prices.at[stock_code, "high"]
            low_price = other_prices.at[stock_code, "low"]

            if trade_status[stock_code] == '交易' and (maxupordown[stock_code] == 0 or (maxupordown[stock_code] != 0 \
            and (open_price != high_price or open_price != low_price))):
                s = self.buy_signal[stock_code]
                stock_name = s[0]
                amount = s[1]
                type = s[-1]
                if amount * open_price * (1 + self.commission) > self.cash:
                    amount = math.floor(self.cash / (1 + self.commission) / open_price / 100) * 100
                if amount > 0:
                    self.cash = self.cash - open_price * amount * (1 + self.commission)
                    # 记录持仓，包括简称、数量、买入价（前复权）、最大浮动收益率、买入类型、日期
                    self.position[stock_code] = [stock_name, amount, open_prices_f[stock_code], 0, type, date]
                    # 记录交易，包括日期、证券代码、交易市场代码、交易方向、交易数量、交易价格
                    tmp = stock_code.split('.')
                    trade_code = tmp[0]
                    market = tmp[1]
                    if market == 'SZ':
                        market = 'XSHE'
                    else:
                        market = 'XSHG'
                    self.transaction.append([date, "09:30:00", trade_code, market, "BUY", '', amount, open_price])
            #无论买信号执行与否，删除买信号
            del self.buy_signal[stock_code]

        # 生成交易记录文件
        writer = pd.ExcelWriter("./" + date + "/交易记录.xls")
        transaction = pd.DataFrame(self.transaction, columns=["日期", "成交时间", "证券代码", "交易市场代码", "交易方向", "投保", "交易数量", "交易价格"])
        transaction.to_excel(writer, "交易记录", index=False)
        writer.save()

    def generateBuySignal(self, date):
        try:
            data = pd.read_excel("data" + str(self.year + 1) + ".xls")
            data.set_index("stock_code", inplace = True)
        except:
            #获取全部A股列表
            sectorconstituent = w.wset("sectorconstituent", "date=" + date + ";sectorid=a001010100000000;field=wind_code,sec_name")
            stock_codes = sectorconstituent.Data[0]
            stock_names = sectorconstituent.Data[1]

            #提取年报预披露时间
            stm_predict_issuingdate = w.wss(stock_codes, "stm_predict_issuingdate","rptDate=" + str(self.year) + "1231").Data[0]

            #提取去年年报披露时间
            stm_issuingdate = w.wss(stock_codes, "stm_issuingdate","rptDate=" + str(self.year - 1) + "1231").Data[0]

            #提取上市时间
            ipo_date = w.wss(stock_codes, "ipo_date").Data[0]

            #提取三季报EPS
            eps_basic = w.wss(stock_codes, "eps_basic","rptDate=" + str(self.year) + "0930;currencyType=").Data[0]

            #提取三季报归属母公司股东净利润同比增长率
            yoynetprofit = w.wss(stock_codes, "yoynetprofit","rptDate=" + str(self.year) + "0930").Data[0]

            data = pd.DataFrame({"stock_code" : stock_codes, "stock_name" : stock_names, \
                                 "stm_predict_issuingdate" : stm_predict_issuingdate, \
                                 "stm_issuingdate" : stm_issuingdate, "ipo_date" : ipo_date, \
                                 "eps_basic" : eps_basic, "yoynetprofit" : yoynetprofit})
            data.set_index("stock_code", inplace=True)
            # 计算年报预披露时间比去年年报实际披露时间提前的天数
            stock_codes = list(data.index)
            n = len(stock_codes)
            days_ahead = [0] * n
            for i in range(n):
                try:
                    days_ahead[i] = (data.at[stock_codes[i], "stm_issuingdate"].replace(year=self.year + 1) - \
                                     data.at[stock_codes[i], "stm_predict_issuingdate"]).days
                except:
                    days_ahead[i] = (data.at[stock_codes[i], "stm_issuingdate"].replace(year=self.year + 1, day=28) - \
                                     data.at[stock_codes[i], "stm_predict_issuingdate"]).days
            data["days_ahead"] = days_ahead
            data = data.ix[:,["stock_name", "stm_predict_issuingdate", "stm_issuingdate", "days_ahead", "ipo_date", "eps_basic", "yoynetprofit"]]

            #将原始数据保存到本地
            writer = pd.ExcelWriter("data" + str(self.year + 1) + ".xls")
            data.to_excel(writer)
            writer.save()

        data.dropna(how = "any", inplace = True)
        #选出年报预披露时间在2月15日之前(含2月15日）,较去年年报实际披露时间提前60日,非次新的股票
        candidates = data[(data.stm_predict_issuingdate <= datetime.datetime(self.year + 1, 2, 15)) & (data.ipo_date < \
                               datetime.datetime(self.year, 1, 1)) & (data.days_ahead >= 60) & (data.days_ahead < 365)]

        #筛选出可能高送转的股票
        high_tran_candidates = self.getHighTranCandidate(candidates)

        #筛选出有ST摘帽预期的股票
        st_stocks = self.getSTStock(candidates)

        #生成买信号
        stock_codes = list(candidates.index)
        n = len(stock_codes)
        close_prices = w.wss(stock_codes, "close", "tradeDate=" + date + ";priceAdj=U;cycle=D").Data[0]
        close_prices = dict(zip(stock_codes, close_prices))
        stock_asset = 1.0 * self.initial_asset_value / n
        for stock_code in stock_codes:
            if stock_code in list(high_tran_candidates.index):
                buy_type = 1
            elif stock_code in list(st_stocks.index):
                buy_type = 2
            else:
                buy_type = 0
            stock_name = candidates.at[stock_code, "stock_name"]
            amount = math.floor(stock_asset / close_prices[stock_code] / 100) * 100
            self.buy_signal[stock_code] = [stock_name, amount, "Buy", buy_type]

        if self.buy_signal:
            xlswr.writeDict2Xls(self.buy_signal, ["代码", "简称", "数量", "方向", "买入类型"], "./" + date + "/买入信号.xls")

        writer = pd.ExcelWriter("股票池.xls")
        candidates.to_excel(writer, "股票池")
        high_tran_candidates.to_excel(writer, "高送转预期股")
        st_stocks.to_excel(writer, "ST摘帽预期股")
        writer.save()

    def generateSellSignal(self, date):
        #无持仓返回
        if not self.position:
            return

        # 提取当天公布的业绩预告
        stock_codes = list(self.position.keys())
        yoynetprofit_forcast = w.wss(stock_codes, "sec_name, profitnotice_changemin, profitnotice_date","rptDate=" + str(self.year) + "1231").Data
        yoynetprofit_forcast = np.vstack((stock_codes, yoynetprofit_forcast)).T
        yoynetprofit_forcast = pd.DataFrame(data=yoynetprofit_forcast,columns=["stock_code", "sec_name","profitnotice_changemin", "profitnotice_date"])
        yoynetprofit_forcast["profitnotice_changemin"] = yoynetprofit_forcast["profitnotice_changemin"].astype(np.float)
        yoynetprofit_forcast.dropna(how='any', inplace=True)
        yoynetprofit_forcast.set_index("stock_code", inplace=True)
        profitnotice_date = datetime.datetime.strptime(date, "%Y%m%d") + datetime.timedelta(days=1)
        yoynetprofit_forcast = yoynetprofit_forcast[yoynetprofit_forcast.profitnotice_date == profitnotice_date]

        # 监控持仓中有高送转预期的股票，如果业绩预告同比增幅<0,则立刻生成卖出信号
        for stock_code in list(yoynetprofit_forcast.index):
            if yoynetprofit_forcast.at[stock_code, "profitnotice_changemin"] < 0 and stock_code in list(self.position.keys()):
                p = self.position[stock_code]
                buy_type = p[-2]
                if buy_type == 1:
                    stock_name = p[0]
                    amount = p[1]
                    sell_type = 0
                    sell_info = "业绩预告归属母公司股东净利润同比增长率小于0"
                    self.sell_signal[stock_code] = [stock_name, amount, "Sell",sell_type, sell_info]

        # 提取当天公布的分红预案
        stock_codes = list(self.position.keys())
        div_plan = w.wss(stock_codes, "sec_name,div_cashbeforetax,div_stock,div_capitalization,div_prelandate,div_preDisclosureDate","rptDate=" + str(self.year) + "1231").Data
        div_plan = np.vstack((stock_codes, div_plan)).T
        div_plan = pd.DataFrame(data = div_plan, columns=["stock_code", "sec_name", "div_cashbeforetax","div_stock","div_capitalization","div_prelandate","div_preDisclosureDate"])
        div_plan["div_cashbeforetax"] = div_plan["div_cashbeforetax"].astype(np.float)
        div_plan["div_stock"] = div_plan["div_stock"].astype(np.float)
        div_plan["div_capitalization"] = div_plan["div_capitalization"].astype(np.float)
        div_plan.dropna(how='any', inplace=True)
        prelandate = datetime.datetime.strptime(date, "%Y%m%d") + datetime.timedelta(days=1)
        div_plan.set_index("stock_code", inplace=True)
        div_plan = div_plan[(div_plan.div_prelandate == prelandate) | (div_plan.div_preDisclosureDate == prelandate)]

        # 监控持仓中有高送转预期的股票，如果公布的分红预案并非高送转，则立刻生成卖出信号
        for stock_code in list(div_plan.index):
            if div_plan.at[stock_code, "div_stock"] + div_plan.at[stock_code, "div_capitalization"] < 0.5 and stock_code in list(self.position.keys()):
                p = self.position[stock_code]
                buy_type = p[-2]
                if buy_type == 1:
                    stock_name = p[0]
                    amount = p[1]
                    sell_type = 1
                    sell_info = "分红预案并没有高送转"
                    self.sell_signal[stock_code] = [stock_name, amount, "Sell",sell_type, sell_info]

        #监控实际发生了高送转的股票，在股权登记日当日生成卖出信号，在除权除息日当天开盘卖出
        rptDate = str(self.year) + "1231"
        dividend = w.wss(list(self.position.keys()), "div_cashbeforetax,div_stock,div_capitalization,div_recorddate",
                         "rptDate=" + rptDate)
        dividend = pd.DataFrame(data=np.matrix(dividend.Data).T, index=dividend.Codes,
                                columns=["div_cashbeforetax", "div_stock", "div_capitalization", "div_recorddate"])
        dividend.dropna(how = 'any', inplace = True)
        dividend = dividend[(dividend.div_recorddate == datetime.datetime.strptime(date, "%Y%m%d")) & \
                            (dividend.div_stock + dividend.div_capitalization > 0.5)]
        for stock_code in list(dividend.index):
            if stock_code in list(self.position.keys()):
                p = self.position[stock_code]
                stock_name = p[0]
                amount = p[1]
                sell_type = 2
                sell_info = "高送转股票股权登记日次日卖出"
                self.sell_signal[stock_code] = [stock_name, amount, "Sell", sell_type, sell_info]

        #监控ST摘帽预期股，在摘帽后恢复交易当日卖出
        stock_codes = list(self.position.keys())
        ST_sectorconstituent = w.wset("sectorconstituent", "date=" + str(self.year) + "-12-31;sectorid=1000006526000000;field=wind_code,sec_name").Data[0]
        st_stocks = list(set(stock_codes) & set(ST_sectorconstituent))
        if st_stocks:
            st_info = w.wss(st_stocks, "sec_name,riskadmonition_date").Data
            st_info = np.vstack((st_stocks, st_info)).T
            st_info = pd.DataFrame(data = st_info, columns = ["stock_code", "sec_name", "riskadmonition_date"])
            n = st_info.shape[0]
            for i in range(n):
                find_date = False
                riskadmonition_date = st_info.at[i, "riskadmonition_date"]
                st_related_date_str = riskadmonition_date.split(",")
                for st_str in st_related_date_str:
                    tmp = st_str.split("：")
                    if "去*ST" in tmp[0] or "去ST" in tmp[0] or "*ST变ST" in tmp[0]:
                        tmp[1] = tmp[1].strip()
                        if tmp[1][0:4] == str(self.year + 1):
                            find_date = True
                            st_info.at[i, "riskadmonition_date"] = tmp[1]
                            break
                if find_date == False:
                    st_info.drop(i, inplace = True)
        else:
            st_info = pd.DataFrame()

        if st_info.empty:
            pass
        else:
            st_info.set_index("stock_code", inplace = True)
            for stock_code in list(st_info.index):
                if stock_code in list(self.position.keys()):
                    riskadmonition_date = datetime.datetime.strptime(str(st_info.at[stock_code, "riskadmonition_date"]), "%Y%m%d") + datetime.timedelta(days=-1)
                    if datetime.datetime.strptime(date, "%Y%m%d") == riskadmonition_date:
                        p = self.position[stock_code]
                        buy_type = p[-2]
                        if buy_type == 2:
                            stock_name = p[0]
                            amount = p[1]
                            sell_type = 3
                            sell_info = "ST股票摘帽复牌当日卖出"
                            self.sell_signal[stock_code] = [stock_name, amount, "Sell", sell_type, sell_info]
        #保存卖出信号到本地
        if self.sell_signal:
            xlswr.writeDict2Xls(self.sell_signal, ["代码", "简称", "数量", "方向", "卖出类型", "备注"], "./" + date + "/卖出信号.xls")

    def generateClearSignal(self, date):
        for stock_code, p in self.position.items():
            stock_name = p[0]
            amount = p[1]
            sell_type = -1
            sell_info = "到期清盘卖出"
            self.sell_signal[stock_code] = [stock_name, amount, "Sell", sell_type, sell_info]
            xlswr.writeDict2Xls(self.sell_signal, ["代码", "简称", "数量", "方向", "卖出类型", "备注"], "./" + date + "/卖出信号.xls")

    def clearInvestCombi(self):
        while len(self.sell_signal) > 0:
            date = w.tdaysoffset(1, self.last_exist_date, "").Data[0][0]
            date = datetime.datetime.strftime(date, '%Y%m%d')
            for stock_code in list(self.sell_signal.keys()):
                trade_info = w.wss(stock_code, "open,high,low,trade_status,maxupordown", "tradeDate=" + date + ";priceAdj=U;cycle=D").Data
                trade_status = trade_info[3][0]
                maxupordown = trade_info[4][0]
                open_price = trade_info[0][0]
                high_price = trade_info[1][0]
                low_price = trade_info[2][0]
                if trade_status == '交易' and (maxupordown == 0 or (maxupordown != 0 and (open_price != low_price or open_price != high_price))):
                    s = self.sell_signal[stock_code]
                    amount = s[1]
                    self.cash = self.cash + open_price * amount * (1 - self.commission)
                    del self.sell_signal[stock_code]
                    del self.position[stock_code]
                    # 记录交易，包括日期、证券代码、交易市场代码、交易方向、交易数量、交易价格
                    tmp = stock_code.split('.')
                    trade_code = tmp[0]
                    market = tmp[1]
                    if market == 'SZ':
                        market = 'XSHE'
                    else:
                        market = 'XSHG'
                    self.transaction.append([date, "09:30:00", trade_code, market, "SELL", '', amount, open_price])

            self.asset_evaluation(date)
            print("Finished process " + date)
            self.last_exist_date = date


    def asset_evaluation(self, date):
        stock_value = 0.0
        stocks_in_position = list(self.position.keys())
        # 按收盘价对组合估值
        n = len(stocks_in_position)
        if n > 0:
            close_prices = w.wss(stocks_in_position, "close", "tradeDate=" + date + ";priceAdj=U;cycle=D").Data[0]
            for i in range(n):
                stock_code = stocks_in_position[i]
                amount = self.position[stock_code][1]
                close_price = close_prices[i]
                stock_value += close_price * amount

        # 处理持仓股分红送转
        rptDate = str(self.year) + "1231"
        self.processDividend(rptDate, date)

        asset_value = stock_value + self.cash
        self.asset_value.append([date, asset_value / self.initial_asset_value, asset_value, self.cash])

        # 生成日终持仓文件
        xlswr.writeDict2Xls(self.position, ["代码", "简称", "数量", "买入价", "最大浮动收益率", "类型", "买入日期"], \
                            "./" + date + "/持仓.xls")

        # 生成日终净值文件
        xlswr.writeList2Xls(self.asset_value, ["日期", "单位净值", "资产净值", "可用现金"], "./" + date + "/净值.xls")

    def processDividend(self, rptDate, date):
        stocks_in_position = list(self.position.keys())
        if len(stocks_in_position) <= 0:
            return

        dividend = w.wss(stocks_in_position, "div_cashbeforetax,div_stock,div_capitalization,div_recorddate","rptDate=" + rptDate)
        dividend = pd.DataFrame(data=np.matrix(dividend.Data).T, index=dividend.Codes,
                                columns=["div_cashbeforetax", "div_stock", "div_capitalization", "div_recorddate"])
        dividend.dropna(how = 'any', inplace = True)

        for stock_code in self.position.keys():
            if stock_code not in dividend.index:
                continue
            d = dividend.loc[stock_code]
            if datetime.datetime.strptime(date, "%Y%m%d") == d["div_recorddate"]:
                p = self.position[stock_code]
                # 确定个人所得税税率
                days_in_position = (d["div_recorddate"] - datetime.datetime.strptime(p[-1], "%Y%m%d")).days
                if days_in_position > 365:
                    tax_ratio = 0.0
                elif days_in_position > 30:
                    tax_ratio = 0.1
                else:
                    tax_ratio = 0.2
                amount = p[1]
                div_cashaftertax = d["div_cashbeforetax"] * amount * (1 - tax_ratio) - d["div_stock"] * tax_ratio * amount
                self.cash += div_cashaftertax
                self.position[stock_code][1] = amount + amount * (d["div_stock"] + d["div_capitalization"])

    def getHighTranCandidate(self, candidates):
        #提取去年三季报每股基本公积，每股留存收益，总股本
        stock_codes = list(candidates.index)
        high_tran_info = w.wss(stock_codes, "sec_name,surpluscapitalps,retainedps,total_shares","rptDate=" + str(self.year) + "0930;unit=1;tradeDate=" + str(self.year) + "1231").Data
        high_tran_info = np.vstack((stock_codes, high_tran_info)).T
        high_tran_info = pd.DataFrame(data = high_tran_info, columns = ["stock_code", "stock_name","surpluscapitalps","retainedps","total_shares"])
        high_tran_info.set_index("stock_code", inplace = True)
        high_tran_info["surpluscapitalps"] = high_tran_info["surpluscapitalps"].astype(np.float)
        high_tran_info["retainedps"] = high_tran_info["retainedps"].astype(np.float)
        high_tran_info["total_shares"] = high_tran_info["total_shares"].astype(np.float)
        #删选出每股资本公积+每股留存收益>3，总股本小于20亿的股票
        high_tran_candidates = high_tran_info[(high_tran_info.surpluscapitalps + high_tran_info.retainedps > 3) & (high_tran_info.total_shares < 20E8)]

        if not high_tran_candidates.empty:
            #提取去年三季报和业绩预告归属母公司股东净利润同比增长率
            stock_codes = list(high_tran_candidates.index)
            yoynetprofit_3rd = w.wss(stock_codes, "yoynetprofit", "rptDate=" + str(self.year) + "0930").Data[0]
            yoynetprofit_3rd = dict(zip(stock_codes, yoynetprofit_3rd))
            yoynetprofit_forcast = w.wss(stock_codes, "profitnotice_changemin, profitnotice_date", "rptDate=" + str(self.year) + "1231").Data
            yoynetprofit_forcast = np.vstack((stock_codes, yoynetprofit_forcast)).T
            yoynetprofit_forcast = pd.DataFrame(data = yoynetprofit_forcast, columns = ["stock_code", "profitnotice_changemin","profitnotice_date"])
            yoynetprofit_forcast.set_index("stock_code", inplace = True)
            yoynetprofit_forcast["profitnotice_changemin"] = yoynetprofit_forcast["profitnotice_changemin"].astype(np.float)
            yoynetprofit_forcast.dropna(how = 'any', inplace = True)
            yoynetprofit_forcast = yoynetprofit_forcast[yoynetprofit_forcast.profitnotice_date <= datetime.datetime(self.year + 1, 1, 1)]

            candidate_list = list(high_tran_candidates.index)
            for stock_code in list(high_tran_candidates.index):
                if stock_code in list(yoynetprofit_forcast.index):
                    if yoynetprofit_forcast.at[stock_code, "profitnotice_changemin"] < 0:
                        candidate_list.remove(stock_code)
                else:
                    if yoynetprofit_3rd[stock_code] < 0:
                        candidate_list.remove(stock_code)
            high_tran_candidates = high_tran_candidates.loc[candidate_list, :]
        return high_tran_candidates

    def getSTStock(self, candidates):
        #找出选出的股票池中的ST股
        ST_sectorconstituent = w.wset("sectorconstituent","date=" + str(self.year) + "-12-31;sectorid=1000006526000000;field=wind_code,sec_name").Data[0]
        st_stocks = []
        for stock_code in list(candidates.index):
            if stock_code in ST_sectorconstituent:
                st_stocks.append(stock_code)
        st_stocks = candidates.loc[st_stocks, :]
        return st_stocks





