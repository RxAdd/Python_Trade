# 先把API com元件初始化
import os

# 第二種讓群益API元件可導入Python code內用的物件宣告
import comtypes.client
from os.path import dirname, abspath, split
#comtypes.client.GetModule( dirname(dirname(abspath(__file__))) + r'\SKCOM.dll' )
import comtypes.gen.SKCOMLib as sk
skC = comtypes.client.CreateObject(sk.SKCenterLib,interface=sk.ISKCenterLib)
skO = comtypes.client.CreateObject(sk.SKOrderLib,interface=sk.ISKOrderLib)

# 畫視窗用物件
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox

# 載入其他物件
import Config
import MessageControl
#----------------------------------------------------------------------------------------------------------------------------------------------------

class FutureOCOOrder(Frame):
    def __init__(self, master=None, information=None):
        Frame.__init__(self)
        self.__master = master
        self.__oMsg = MessageControl.MessageControl()
        # UI variable
        self.__dOrder = dict(
            listInformation = information,
            txtID = '',
            boxAccount = '',
        )

        self.__CreateWidget()
        
    def SetID(self, id):
        self.__dOrder['txtID'] = id    

    def SetAccount(self, account):
        self.__dOrder['boxAccount'] = account

    def __CreateWidget(self):
        group = LabelFrame(self.__master, text="期貨OCO委託", style="Pink.TLabelframe")
        group.grid(column = 0, row = 3, padx = 10, pady = 10)

        frame = Frame(group, style="Pink.TFrame")
        frame.grid(column = 0, row = 0, padx = 10, pady = 5, sticky = 'ew')
        frame.grid_columnconfigure(7, minsize = 10)

        # 商品代碼
        lbStockNo = Label(frame, style="Pink.TLabel", text = "商品代碼")
        lbStockNo.grid(column = 0, row = 0, pady = 3)
            # 輸入框
        txtStockNo = Entry(frame, width = 10)
        txtStockNo.grid(column = 0, row = 1, rowspan = 2, padx = 10, pady = 3)

        # 買賣別
        lbBuySell = Label(frame, style="Pink.TLabel", text = "買賣別")
        lbBuySell.grid(column = 1, row = 0)
            # 輸入框
        boxBuySell = Combobox(frame, width = 5, state='readonly')
        boxBuySell['values'] = Config.BUYSELLSET
        boxBuySell.grid(column = 1, row = 1, padx = 10)

        # 買賣別2
            # 輸入框
        boxBuySell2 = Combobox(frame, width = 5, state='readonly')
        boxBuySell2['values'] = Config.BUYSELLSET
        boxBuySell2.grid(column = 1, row = 2, padx = 10)

        # 委託條件
        lbPeriod = Label(frame, style="Pink.TLabel", text = "委託條件")
        lbPeriod.grid(column = 2, row = 0)
            # 輸入框
        boxPeriod = Combobox(frame, width = 8, state='readonly')
        boxPeriod['values'] = Config.PERIODSET['future']
        boxPeriod.grid(column = 2, row = 1, rowspan = 2, padx = 10)

        # 倉別
        lbNewClose = Label(frame, style="Pink.TLabel", text = "倉別")
        lbNewClose.grid(column = 3, row = 0)
            # 輸入框
        boxNewClose = Combobox(frame, width = 5, state='readonly')
        boxNewClose['values'] = Config.NEWCLOSESET['future']
        boxNewClose.grid(column = 3, row = 1, rowspan = 2, padx = 10)

        # 當沖與否
        lbFlag = Label(frame, style="Pink.TLabel", text = "當沖與否")
        lbFlag.grid(column = 4, row = 0)
            # 輸入框
        boxFlag = Combobox(frame, width = 8, state='readonly')
        boxFlag['values'] = Config.FLAGSET['future']
        boxFlag.grid(column = 4, row = 1, rowspan = 2, padx = 10)

        # 委託價
        lbPrice = Label(frame, style="Pink.TLabel", text = "委託價")
        lbPrice.grid(column = 5, row = 0)
            # 輸入框
        txtPrice = Entry(frame, width = 10)
        txtPrice.grid(column = 5, row = 1, padx = 10)

        # 委託價2
            # 輸入框
        txtPrice2 = Entry(frame, width = 10)
        txtPrice2.grid(column = 5, row = 2, padx = 10)

        # 委託量
        lbQty = Label(frame, style="Pink.TLabel", text = "委託量")
        lbQty.grid(column = 6, row = 0)
            # 輸入框
        txtQty = Entry(frame, width = 10)
        txtQty.grid(column = 6, row = 1, rowspan = 2, padx = 10)

        # 觸發價
        Label(frame, style="Pink.TLabel", text = ">=").grid(column = 7, row = 1)
        lbTrigger = Label(frame, style="Pink.TLabel", text = "觸發價")
        lbTrigger.grid(column = 8, row = 0)
            # 輸入框
        txtTrigger = Entry(frame, width = 10)
        txtTrigger.grid(column = 8, row = 1, padx = 10)

        # 觸發價2
        Label(frame, style="Pink.TLabel", text = "<=").grid(column = 7, row = 2)
            # 輸入框
        txtTrigger2 = Entry(frame, width = 10)
        txtTrigger2.grid(column = 8, row = 2, padx = 10)

        # 盤別
        lbReserved = Label(frame, style="Pink.TLabel", text = "盤別")
        lbReserved.grid(column = 9, row = 0)
            # 輸入框
        boxReserved = Combobox(frame, width = 8, state='readonly')
        boxReserved['values'] = Config.RESERVEDSET
        boxReserved.grid(column = 9, row = 1, rowspan = 2, padx = 10)

        # btnSendOrder
        btnSendOrder = Button(frame, style = "Pink.TButton", text = "送出委託")
        btnSendOrder["command"] = self.__btnSendOrder_Click
        btnSendOrder.grid(column = 10, row =  1, padx = 10)

        self.__dOrder['txtStockNo'] = txtStockNo
        self.__dOrder['boxBuySell'] = boxBuySell
        self.__dOrder['boxBuySell2'] = boxBuySell2
        self.__dOrder['boxPeriod'] = boxPeriod
        self.__dOrder['boxNewClose'] = boxNewClose
        self.__dOrder['boxFlag'] = boxFlag
        self.__dOrder['txtPrice'] = txtPrice
        self.__dOrder['txtPrice2'] = txtPrice2
        self.__dOrder['txtQty'] = txtQty
        self.__dOrder['txtTrigger'] = txtTrigger
        self.__dOrder['txtTrigger2'] = txtTrigger2
        self.__dOrder['boxReserved'] = boxReserved

    # 4.下單送出
    # sBuySell, sBuySell2, sTradeType, sNewClose, sDayTrade, sReserved
    def __btnSendOrder_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇期貨帳號！')            
        else:
            self.__SendOrder_Click(False)

    def __SendOrder_Click(self, bAsyncOrder):
        try:            
            if self.__dOrder['boxBuySell'].get() == "買進":
                sBuySell = 0
            elif self.__dOrder['boxBuySell'].get() == "賣出":
                sBuySell = 1

            if self.__dOrder['boxBuySell2'].get() == "買進":
                sBuySell2 = 0
            elif self.__dOrder['boxBuySell2'].get() == "賣出":
                sBuySell2 = 1

            if self.__dOrder['boxPeriod'].get() == "ROD":
                sTradeType = 0
            elif self.__dOrder['boxPeriod'].get() == "IOC":
                sTradeType = 1
            elif self.__dOrder['boxPeriod'].get() == "FOK":
                sTradeType = 2

            if self.__dOrder['boxNewClose'].get() == "新倉":
                sNewClose = 0
            elif self.__dOrder['boxNewClose'].get() == "平倉":
                sNewClose = 1
            elif self.__dOrder['boxNewClose'].get() == "自動":
                sNewClose = 2
            
            if self.__dOrder['boxFlag'].get() == "非當沖":
                sDayTrade = 0
            elif self.__dOrder['boxFlag'].get() == "當沖":
                sDayTrade = 1
            
            if self.__dOrder['boxReserved'].get() == "盤中":
                sReserved = 0
            elif self.__dOrder['boxReserved'].get() == "T盤預約":
                sReserved = 1

            # 建立下單用的參數(FUTUREOCOORDER)物件(下單時要填商品代號,買賣別,委託價,數量等等的一個物件)
            oOrder = sk.FUTUREOCOORDER()
            # 填入完整帳號
            oOrder.bstrFullAccount =  self.__dOrder['boxAccount']
            # 填入期權代號
            oOrder.bstrStockNo = self.__dOrder['txtStockNo'].get()
            # 買賣別
            oOrder.sBuySell = sBuySell
            oOrder.sBuySell2 = sBuySell2
            # ROD、IOC、FOK
            oOrder.sTradeType = sTradeType
            # 新倉、平倉、自動
            oOrder.sNewClose = sNewClose
            # 非當沖、當沖
            oOrder.sDayTrade = sDayTrade
            # 委託價
            oOrder.bstrPrice = self.__dOrder['txtPrice'].get()
            oOrder.bstrPrice2 = self.__dOrder['txtPrice2'].get()
            # 委託數量
            oOrder.nQty = int(self.__dOrder['txtQty'].get())
            # 觸發價
            oOrder.bstrTrigger = self.__dOrder['txtTrigger'].get()
            oOrder.bstrTrigger2 = self.__dOrder['txtTrigger2'].get()
            # 盤中、T盤預約
            oOrder.sReserved = sReserved

            message, m_nCode = skO.SendFutureOCOOrder(self.__dOrder['txtID'], bAsyncOrder, oOrder)
            self.__oMsg.SendReturnMessage("Order", m_nCode, "SendFutureOCOOrder", self.__dOrder['listInformation'])
        except Exception as e:
            messagebox.showerror("error！", e)
