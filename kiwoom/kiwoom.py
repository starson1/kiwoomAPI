from PyQt5.QAxContainer import *
from config.errCode import *
from PyQt5.QtCore import *
from PyQt5.QtTest import *

class Kiwoom(QAxWidget):

    def __init__(self):
        super().__init__()
        ################eventloop####################
        self.login_event_loop = QEventLoop()
        self.detail_account_info_event_loop=QEventLoop()
        self.detail_account_info_event_loop = QEventLoop()
        self.calculator_event_loop= QEventLoop()
        #############################################
        ################Screen number################
        self.screen_my_info = "2000"
        self.screen_calculation_stock= "4000"
        #############################################
        ################variables####################
        self.account_num = None
        self.account_stock_dict={}
        self.not_account_stock_dict={}
        #############################################

        ################계좌관련변수##################
        self.use_money =0
        self.use_money_per = 0.5
        #############################################

        ################종목분석용##################
        self.cal_data = []
        #############################################

        self.get_ocx_instance()
        self.event_slots()

        self.signal_login_commConnect() # 로그인하기
        self.get_account_info() # 계좌정보 가져오기
        self.detail_account_info() # 예수금 정보 가져오기
        self.detail_account_mystock()  # 계좌 잔고 내역 요청
        self.not_traded() # 미체결 요청

        self.calculator_fnc() #

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)
    def login_slot(self, errCode):
        print(errCode)
        print(errors(errCode))

        self.login_event_loop.exit()
    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode,sRecordName, sPrevNext):
        '''
        tr 요청받는 slot
        :param sScrNo: 스크린 번호
        :param sRQName: 요청시 만든 이름
        :param sTrCode: 요청 tr 코드
        :param sRecordName: 사용 안함..!
        :param sPrevNext: 다음 페이지 있는지 여부(연속조회 유무)
        :return:
        '''

        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String,String,int,String)",sTrCode,sRQName,0,"예수금")
            print("예수금 %d" % int(deposit))
            deposit = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode, sRQName, 0, "출금가능금액")
            print("출금가능금액 %d" % int(deposit))

            self.use_money = int(deposit) * self.use_money_per
            self.use_money /=4

            #event loop 탈출
            self.detail_account_info_event_loop.exit()
        elif sRQName =="계좌평가잔고내역요청":
            print("계좌평가잔고내역요청")

            total_buy_money = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0,"총매입금액")
            self.total_buy_money = int(total_buy_money.strip())
            total_profit_loss_money = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,0, "총평가손익금액")
            self.total_profit_loss_money = int(total_profit_loss_money.strip())
            total_profit_loss_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,0, "총수익률(%)")

            self.total_profit_loss_rate = float(total_profit_loss_rate.strip())

            print("계좌평가잔고내역요청 싱글데이터 : %s - %s - %s" % (total_buy_money, total_profit_loss_money, total_profit_loss_rate))

            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"종목번호")  # 출력 : A039423 // 알파벳 A는 장내주식, J는 ELW종목, Q는 ETN종목
                code = code.strip()[1:]

                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"종목명")  # 출럭 : 한국기업평가
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"보유수량")  # 보유수량 : 000000000000010
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"매입가")  # 매입가 : 000000000054100
                earn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"수익률(%)")  # 수익률 : -000000001.94
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"현재가")  # 현재가 : 000000003450
                total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"매매가능수량")



                if code in self.account_stock_dict:  # dictionary 에 해당 종목이 있나 확인
                    pass
                else:
                    self.account_stock_dict[code] = {}

                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                earn_rate = float(earn_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())

                self.account_stock_dict[code].update({"종목명": code_nm})
                self.account_stock_dict[code].update({"보유수량": stock_quantity})
                self.account_stock_dict[code].update({"매입가": buy_price})
                self.account_stock_dict[code].update({"수익률(%)": earn_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"매입금액": total_chegual_price})
                self.account_stock_dict[code].update({'매매가능수량': possible_quantity})
                print("종목코드: %s - 종목명: %s - 보유수량: %s - 매입가:%s - 수익률: %s - 현재가: %s" % (code, code_nm, stock_quantity, buy_price, earn_rate, current_price))

            print("sPreNext : %s" % sPrevNext)
            print("계좌에 가지고 있는 종목은 %s " % rows)

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()
        elif sRQName == "실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString,QString)", sTrCode, sRQName)  # 보유종목수 count, 종목별로 정보 받아옴
            for i in range(rows):
                code = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode, sRQName, i, "종목번호")
                code_name = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode, sRQName, i, "종목명")
                order_no = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode, sRQName, i, "주문번호")
                order_status = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode, sRQName, i, "주문상태") # 접수/확인/체결
                order_quantity = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode, sRQName, i, "주문수량")
                order_price = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode, sRQName, i, "주문가격")
                order_category = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode, sRQName, i, "주문구분")
                not_traded = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode, sRQName, i, "미체결수량")
                yes_traded = self.dynamicCall("GetCommData(String,String,int,String)", sTrCode, sRQName, i, "체결량")

                code = code.strip()
                code_name = code_name.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_category = order_category.strip().lstrip('+').lstrip('-')
                not_traded = int(not_traded.strip())
                yes_traded = int(yes_traded.strip())

                if order_no in self.not_account_stock_dict:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}

                self.not_account_stock_dict[order_no].update({"종목코드":code})
                self.not_account_stock_dict[order_no].update({"종목명": code_name})
                self.not_account_stock_dict[order_no].update({"주문번호": order_no})
                self.not_account_stock_dict[order_no].update({"주문상태": order_status})
                self.not_account_stock_dict[order_no].update({"주문수량": order_quantity})
                self.not_account_stock_dict[order_no].update({"주문가격": order_price})
                self.not_account_stock_dict[order_no].update({"주문구분": order_category})
                self.not_account_stock_dict[order_no].update({"미체결수량": not_traded})
                self.not_account_stock_dict[order_no].update({"체결량": yes_traded})

                print("미체결종목 %s" % self.not_account_stock_dict[order_no])

            self.detail_account_info_event_loop.exit()
        elif  sRQName == "주식일봉차트조회":
            code = self.dynamicCall("GetCommData(QString,QString,int,QString)", sTrCode, sRQName, 0, "종목코드")
            code = code.strip()
            print("%s 일봉 데이터 요청" % code)

            cnt = self.dynamicCall("GetRepeatCnt(QString,QString)", sTrCode, sRQName)
            print("number of days : %s" % cnt)

            for i in range(cnt):
                data =[]
                current_price = self.dynamicCall("GetCommData(QString,QString,int,QString)",sTrCode,sRQName,i,"현재가")
                value = self.dynamicCall("GetCommData(QString,QString,int,QString)", sTrCode, sRQName, i, "거래량")
                trading_value = self.dynamicCall("GetCommData(QString,QString,int,QString)", sTrCode, sRQName, i, "거래대금")
                date = self.dynamicCall("GetCommData(QString,QString,int,QString)", sTrCode, sRQName, i, "일자")
                high_price = self.dynamicCall("GetCommData(QString,QString,int,QString)", sTrCode, sRQName, i, "고가")
                low_price = self.dynamicCall("GetCommData(QString,QString,int,QString)", sTrCode, sRQName, i, "저가")
                start_price = self.dynamicCall("GetCommData(QString,QString,int,QString)", sTrCode, sRQName, i, "시가")

                data.append("")
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_value.strip())
                data.append(date.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                data.append(start_price.strip())
                data.append("")

                self.cal_data.append(data.copy())

            print(self.cal_data)

            if sPrevNext == "2":
                self.chart_view(code=code,sPrevNext=sPrevNext)
            else:
                print("총 일수 %s "% len(self.cal_data))
                if self.cal_data == None or len(self.cal_data)<120:
                    pass_success = False
                else:
                    #120 이평선 구하기
                    total_price =0
                    for value in self.cal_data[:120]:
                        total_price += int(value[1])
                    avg120_price = total_price / 120


                    bottom_stock_price = False
                    chk_price = None
                    if int(self.cal_data[0][7])<= avg120_price:
                        if avg120_price >=int(self.cal_data[0][6]):
                            print("오늘 주가 120 이평선 조건 충족!")
                            bottom_stock_price = True
                            chk_price = int(self.cal_data[0][6])

                    #과거 일봉들이 120 이평선 밑에 있는지 확인.
                    prev_price = None # 과거의 일봉 저가
                    if bottom_stock_price == True:
                        avg120_price = 0
                        price_top = False

                        idx =1
                        while True:
                            if len(self.cal_data[idx:]) < 120:
                                print("NO 120 data!")
                                break
                            total_price =0
                            for value in self.cal_data[idx:120+idx]:
                                total_price == int(value[1])
                            avg120_price_prev = total_price / 120

                            if avg120_price_prev <= int(self.cal_data[idx][6]):
                                if idx <= 20:
                                    print("20일 동안 120 이평선과 같거나 밑에 있으면 조건 충족 x ")
                                    price_top = False
                                    break
                            elif int(self.cal_data[idx][7]) > avg120_price_prev:
                                if idx > 20:
                                    print("120 이평선 위에 있는 일봉 확인됨.")
                                    price_top = True
                                    prev_price = int(self.cal_data[idx][7])
                                    break

                            idx +=1
                        if price_top == True:
                            if avg120_price > avg120_price_prev:
                                if chk_price > prev_price:
                                    print("!!FOUND!!")
                                    pass_success = True
                if pass_success == True:
                    print("qualified")
                    code_nm = self.dynamicCall("GetMasterCodeName(QString)",code)
                    f = open("files/condition_stock.txt","a", encoding='utf-8')
                    f.write("%s\t%s\t%s\n"% (code,code_nm,str(self.cal_data[0][1])))
                    f.close()
                elif pass_success == False:
                    print("not qualified")

                self.cal_data.clear()
                self.calculator_event_loop.exit()



    #TRdata
    def get_account_info(self):
        print("계좌 정보 요청")
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.account_num = account_list.split(';')[0]
        print("my account num : %s" % self.account_num) #8150840811
    def detail_account_info(self):
        print("예수금 정보 요청")
        self.dynamicCall("SetInputValue(String,String)","계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String,String)", "비밀번호", '0000')
        self.dynamicCall("SetInputValue(String,String)", "비밀번호 입력매체구분", '00')
        self.dynamicCall("SetInputValue(String,String)", "조회구분", '2')
        self.dynamicCall("CommRqData(String,String,int,String)","예수금상세현황요청","opw00001","0",self.screen_my_info)

        #event loop
        self.detail_account_info_event_loop.exec_()
    def detail_account_mystock(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가잔고내역요청", "opw00018", sPrevNext,self.screen_my_info)
        # event loop
        self.detail_account_info_event_loop.exec_()
    def not_traded(self,sPrevNext="0"):
        print("미체결 요청")
        self.dynamicCall("SetInputValue(String,String)","계좌번호",self.account_num)
        self.dynamicCall("SetInputValue(String,String)", "체결구분", "1")
        self.dynamicCall("SetInputValue(String,String)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString,QString,int,Qstring)","실시간미체결요청","opt10075",sPrevNext,self.screen_my_info)

        self.detail_account_info_event_loop.exec_()
    def get_code_list(self,market_code):
        code_list = self.dynamicCall("GetCodeListByMarket(QString",market_code)
        code_list = code_list.split(";")[:-1]
        return code_list
    def calculator_fnc(self):
        code_list = self.get_code_list("10")
        print("코스닥 갯수 %s" % len(code_list))

        for idx,code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(QString)",self.screen_calculation_stock)
            print("%s / %s : KOSDAQ Stock Code : %s is updating " % (idx+1,len(code_list),code))
            self.chart_view(code=code)
    def chart_view(self,code=None,date=None,sPrevNext="0"):
        QTest.qWait(3600)
        self.dynamicCall("SetInputValue(QString,QString)","종목코드",code)
        if date != None:
            self.dynamicCall("SetInputValue(QString,QString)", "기준일자",date)
        self.dynamicCall("SetInputValue(QString,QString)", "수정주가구분", "1")
        self.dynamicCall("CommRqData(Qstring,QString,int, QString)","주식일봉차트조회","opt10081",sPrevNext,self.screen_calculation_stock)
        self.calculator_event_loop.exec_()



