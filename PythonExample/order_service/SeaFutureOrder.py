# 先把API com元件初始化
import os

# 第二種讓群益API元件可導入Python code內用的物件宣告
import comtypes.client
#comtypes.client.GetModule(os.path.split(os.path.realpath(__file__))[0] + r'\SKCOM.dll')
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

class Order(Frame):
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
        group = LabelFrame(self.__master, text="海期委託", style="Pink.TLabelframe")
        group.grid(column = 0, row = 0, padx = 10, pady = 10)

        frame = Frame(group, style="Pink.TFrame")
        frame.grid(column = 0, row = 0, padx = 10, pady = 5, sticky = 'ew')

        # 交易所代號
        lbExchangeNo = Label(frame, style="Pink.TLabel", text = "交易所代號")
        lbExchangeNo.grid(column = 0, row = 0, pady = 3)
            # 輸入框
        txtExchangeNo = Entry(frame, width = 10)
        txtExchangeNo.grid(column = 0, row = 1, padx = 10, pady = 3)

        # 商品代碼
        lbStockNo = Label(frame, style="Pink.TLabel", text = "商品代碼")
        lbStockNo.grid(column = 1, row = 0, pady = 3)
            # 輸入框
        txtStockNo = Entry(frame, width = 10)
        txtStockNo.grid(column = 1, row = 1, padx = 10, pady = 3)

        def __clear_entry(event, txtYearMonth):
            txtYearMonth.delete(0, END)
            txtYearMonth['foreground'] = "#000000"
        # 商品年月
        lbYearMonth = Label(frame, style="Pink.TLabel", text = "商品年月")
        lbYearMonth.grid(column = 2, row = 0, pady = 3)
            # 輸入框
        txtYearMonth = Entry(frame, width = 10)
        txtYearMonth.grid(column = 2, row = 1, padx = 10, pady = 3)
        txtYearMonth['foreground'] = "#b3b3b3"
        txtYearMonth.insert(0, "YYYYMM")
        txtYearMonth.bind("<FocusIn>", lambda event: __clear_entry(event, txtYearMonth))

        # 委託價
        lbOrder = Label(frame, style="Pink.TLabel", text = "委託價")
        lbOrder.grid(column = 3, row = 0)
            # 輸入框
        txtOrder = Entry(frame, width = 10)
        txtOrder.grid(column = 3, row = 1, padx = 10)

        # 委託價分子
        lbOrderNumerator = Label(frame, style="Pink.TLabel", text = "委託價分子")
        lbOrderNumerator.grid(column = 4, row = 0)
            # 輸入框
        txtOrderNumerator = Entry(frame, width = 10)
        txtOrderNumerator.grid(column = 4, row = 1, padx = 10)

        # 觸發價
        lbTrigger = Label(frame, style="Pink.TLabel", text = "觸發價")
        lbTrigger.grid(column = 5, row = 0)
            # 輸入框
        txtTrigger = Entry(frame, width = 10)
        txtTrigger.grid(column = 5, row = 1, padx = 10)

        # 觸發價分子
        lbTriggerNumerator = Label(frame, style="Pink.TLabel", text = "觸發價分子")
        lbTriggerNumerator.grid(column = 6, row = 0)
            # 輸入框
        txtTriggerNumerator = Entry(frame, width = 10)
        txtTriggerNumerator.grid(column = 6, row = 1, padx = 10)

        # 委託量
        lbQty = Label(frame, style="Pink.TLabel", text = "委託量")
        lbQty.grid(column = 7, row = 0)
            # 輸入框
        txtQty = Entry(frame, width = 10)
        txtQty.grid(column = 7, row = 1, padx = 10)
        #--------------------------------------------------------------------------------------------------------------------------------------------
    
        frame = Frame(group, style="Pink.TFrame")
        frame.grid(column = 0, row = 1, padx = 10, pady = 5, sticky = 'ew')

        # 買賣別
        lbBuySell = Label(frame, style="Pink.TLabel", text = "買賣別")
        lbBuySell.grid(column = 0, row = 2)
            # 輸入框
        boxBuySell = Combobox(frame, width = 10, state='readonly')
        boxBuySell['values'] = Config.BUYSELLSET
        boxBuySell.grid(column = 0, row = 3, padx = 10)

        # 倉別
        lbNewClose = Label(frame, style="Pink.TLabel", text = "倉別")
        lbNewClose.grid(column = 1, row = 2)
            # 輸入框
        boxNewClose = Combobox(frame, width = 10, state='readonly')
        boxNewClose['values'] = Config.NEWCLOSESET['sea_future']
        boxNewClose.grid(column = 1, row = 3, padx = 10)

        # 當沖與否
        lbFlag = Label(frame, style="Pink.TLabel", text = "當沖與否")
        lbFlag.grid(column = 2, row = 2)
            # 輸入框
        boxFlag = Combobox(frame, width = 10, state='readonly')
        boxFlag['values'] = Config.FLAGSET['future']
        boxFlag.grid(column = 2, row = 3, padx = 10)

        # 委託條件
        lbPeriod = Label(frame, style="Pink.TLabel", text = "委託條件")
        lbPeriod.grid(column = 3, row = 2)
            # 輸入框
        boxPeriod = Combobox(frame, width = 10, state='readonly')
        boxPeriod['values'] = Config.PERIODSET['sea_future']
        boxPeriod.grid(column = 3, row = 3, padx = 10)

        # 委託類型
        lbSpecialTradeType = Label(frame, style="Pink.TLabel", text = "委託類型")
        lbSpecialTradeType.grid(column = 4, row = 2, columnspan = 2)
            # 輸入框
        boxSpecialTradeType = Combobox(frame, width = 20, state='readonly')
        boxSpecialTradeType['values'] = Config.STRADETYPE['sea_future']
        boxSpecialTradeType.grid(column = 4, row = 3, padx = 10, columnspan = 2)

        # btnSendOrder
        btnSendOrder = Button(frame, style = "Pink.TButton", text = "送出委託")
        btnSendOrder["command"] = self.__btnSendOrder_Click
        btnSendOrder.grid(column = 8, row =  3, padx = 10)

        self.__dOrder['txtExchangeNo'] = txtExchangeNo
        self.__dOrder['txtStockNo'] = txtStockNo
        self.__dOrder['txtYearMonth'] = txtYearMonth
        self.__dOrder['txtOrder'] = txtOrder
        self.__dOrder['txtOrderNumerator'] = txtOrderNumerator
        self.__dOrder['txtTrigger'] = txtTrigger
        self.__dOrder['txtTriggerNumerator'] = txtTriggerNumerator
        self.__dOrder['txtQty'] = txtQty

        self.__dOrder['boxBuySell'] = boxBuySell
        self.__dOrder['boxNewClose'] = boxNewClose
        self.__dOrder['boxFlag'] = boxFlag
        self.__dOrder['boxPeriod'] = boxPeriod
        self.__dOrder['boxSpecialTradeType'] = boxSpecialTradeType
    
    # 4.下單送出
    # sBuySell, sNewClose, boxFlag, boxPeriod, sSpecialTradeType
    def __btnSendOrder_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇海期帳號！')            
        else:
            self.__SendOrder_Click(False)

    def __SendOrder_Click(self, bAsyncOrder):
        try:            
            if self.__dOrder['boxBuySell'].get() == "買進":
                sBuySell = 0
            elif self.__dOrder['boxBuySell'].get() == "賣出":
                sBuySell = 1

            if self.__dOrder['boxNewClose'].get() == "新倉":
                sNewClose = 0

            if self.__dOrder['boxFlag'].get() == "非當沖":
                sDayTrade = 0
            elif self.__dOrder['boxFlag'].get() == "當沖":
                sDayTrade = 1

            if self.__dOrder['boxPeriod'].get() == "ROD":
                sTradeType = 0
            
            if self.__dOrder['boxSpecialTradeType'].get() == "LMT（限價）":
                sSpecialTradeType = 0
            elif self.__dOrder['boxSpecialTradeType'].get() == "MKT（市價）":
                sSpecialTradeType = 1
            elif self.__dOrder['boxSpecialTradeType'].get() == "STL（停損限價）":
                sSpecialTradeType = 2
            elif self.__dOrder['boxSpecialTradeType'].get() == "STP（停損市價）":
                sSpecialTradeType = 3

            # 建立下單用的參數(OVERSEAFUTUREORDER)物件(下單時要填商品代號,買賣別,委託價,數量等等的一個物件)
            oOrder = sk.OVERSEAFUTUREORDER()
            # 填入完整帳號
            oOrder.bstrFullAccount =  self.__dOrder['boxAccount']
            # 填入交易所代號
            oOrder.bstrExchangeNo = self.__dOrder['txtExchangeNo'] .get()
            # 填入期權代號
            oOrder.bstrStockNo = self.__dOrder['txtStockNo'].get()
            # 近月商品年月
            oOrder.bstrYearMonth = self.__dOrder['txtYearMonth'].get()
            # 委託價
            oOrder.bstrOrder = self.__dOrder['txtOrder'].get()
            # 委託價分子
            oOrder.bstrOrderNumerator = self.__dOrder['txtOrderNumerator'].get()
            # 觸發價
            oOrder.bstrTrigger = self.__dOrder['txtTrigger'].get()
            # 觸發價分子
            oOrder.bstrTriggerNumerator = self.__dOrder['txtTriggerNumerator'].get()
            # 委託數量
            oOrder.nQty = int(self.__dOrder['txtQty'].get())

            # 買賣別
            oOrder.sBuySell = sBuySell
            # 新倉
            oOrder.sNewClose = sNewClose
            # 非當沖、當沖
            oOrder.sDayTrade = sDayTrade
            # ROD
            oOrder.sTradeType = sTradeType
            # LMT（限價）、MKT（市價）、STL（停損限價）、STP（停損市價）
            oOrder.sSpecialTradeType = sSpecialTradeType

            message, m_nCode = skO.SendOverSeaFutureOrder(self.__dOrder['txtID'], bAsyncOrder, oOrder)
            self.__oMsg.SendReturnMessage("Order", m_nCode, "SendOverSeaFutureOrder", self.__dOrder['listInformation'])
        except Exception as e:
            messagebox.showerror("error！", e)

class DecreaseOrder():
    def __init__(self, master=None, information=None):
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
        group = LabelFrame(self.__master, text="委託減量", style="Pink.TLabelframe")
        group.grid(column = 0, row = 1, padx = 10, pady = 10, sticky = 'ew')
        
        frame = Frame(group, style="Pink.TFrame")
        frame.grid(column = 0, row = 0, padx = 10, pady = 5)

        # 委託序號
        Label(frame, style="Pink.TLabel", text = "委託序號").grid(column = 0, row = 0, pady = 3)
            # 輸入框
        txtSqlNo = Entry(frame, width = 18)
        txtSqlNo.grid(column = 1, row = 0, padx = 10)

        Label(frame, style="PinkFiller.TLabel", text = "一一").grid(column = 2, row = 0)

        # 欲減少的數量
        Label(frame, style="Pink.TLabel", text = "欲減少的數量").grid(column = 3, row = 0)
            # 輸入框
        txtDecreaseQty = Entry(frame, width = 10)
        txtDecreaseQty.grid(column = 4, row = 0, padx = 10)

        Label(frame, style="PinkFiller.TLabel", text = "一一").grid(column = 5, row = 0)

        # btnSendOrder
        btnSendOrder = Button(frame, style = "Pink.TButton", text = "送出")
        btnSendOrder["command"] = self.__btnSendOrder_Click
        btnSendOrder.grid(column = 6, row =  0, padx = 10)

        self.__dOrder['txtSqlNo'] = txtSqlNo
        self.__dOrder['txtDecreaseQty'] = txtDecreaseQty

    # 送出
    def __btnSendOrder_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇海期帳號！')            
        else:
            self.__SendOrder_Click(False)

    def __btnSendOrderAsync_Click(self):
        self.__SendOrder_Click(True)

    def __SendOrder_Click(self, bAsyncOrder):
        try:
            message, m_nCode = skO.OverSeaDecreaseOrderBySeqNo( self.__dOrder['txtID'], bAsyncOrder, self.__dOrder['boxAccount'],\
                self.__dOrder['txtSqlNo'].get(), int(self.__dOrder['txtDecreaseQty'].get()) )
            self.__oMsg.SendReturnMessage("Order", m_nCode, "OverSeaDecreaseOrderBySeqNo", self.__dOrder['listInformation'])
        except Exception as e:
            messagebox.showerror("error！", e)

class CancelOrder():
    def __init__(self, master=None, information=None):
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
        group = LabelFrame(self.__master, text="取消委託", style="Pink.TLabelframe")
        group.grid(column = 0, row = 2, padx = 10, pady = 10, sticky = 'ew')
        
        frame = Frame(group, style="Pink.TFrame")
        frame.grid(column = 0, row = 0, padx = 10, pady = 5)

        # 委託序號、委託書號
        row = 0
        self.__radVar = IntVar()

        for _ in 'txtSeqNo', 'txtBookNo':
                # 輸入框
            self.__dOrder[_] = Entry(frame, width = 18, state = 'disabled')
            self.__dOrder[_].grid(column = 1, row = row, padx = 10)

            if _ == 'txtSeqNo':
                text = '委託序號'
            elif _ == 'txtBookNo':
                text = '委託書號'

            rb = Radiobutton(frame, style="Pink.TRadiobutton", text = text, variable = self.__radVar, value = row, command = self.__radCall)
            rb.grid(column = 0, row = row, pady = 3, sticky = 'ew')

            row = row + 1
        self.__dOrder['txtSeqNo']['state'] = '!disabled'

        # btnSendOrder
        btnSendOrder = Button(frame, style = "Pink.TButton", text = "送出")
        btnSendOrder["command"] = self.__btnSendOrder_Click
        btnSendOrder.grid(column = 2, row =  1, padx = 50)

    def __radCall(self):
        radSel = self.__radVar.get()
        self.__dOrder['txtSeqNo']['state'] = '!disabled' if radSel == 0 else 'disabled'
        self.__dOrder['txtBookNo']['state'] = '!disabled' if radSel == 1 else 'disabled'

    # 送出
    def __btnSendOrder_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇海期帳號！')            
        elif self.__radVar.get() == 0 and self.__dOrder['txtSeqNo'].get() == '':
            messagebox.showerror("error！", '請輸入欲取消的委託序號！')    
        elif self.__radVar.get() == 1 and self.__dOrder['txtBookNo'].get() == '':
            messagebox.showerror("error！", '請輸入欲取消的委託書號！')            
        else:
            self.__SendOrder_Click(False)

    def __SendOrder_Click(self, bAsyncOrder):
        try:
            if self.__radVar.get() == 0:
                message, m_nCode = skO.OverSeaCancelOrderBySeqNo( self.__dOrder['txtID'], bAsyncOrder, self.__dOrder['boxAccount'],\
                    self.__dOrder['txtSeqNo'].get() )
                self.__oMsg.SendReturnMessage("Order", m_nCode, "OverSeaCancelOrderBySeqNo", self.__dOrder['listInformation'])
            elif self.__radVar.get() == 1:
                message, m_nCode = skO.OverSeaCancelOrderByBookNo( self.__dOrder['txtID'], bAsyncOrder, self.__dOrder['boxAccount'],\
                    self.__dOrder['txtBookNo'].get() )
                self.__oMsg.SendReturnMessage("Order", m_nCode, "OverSeaCancelOrderByBookNo", self.__dOrder['listInformation'])
        except Exception as e:
            messagebox.showerror("error！", e)

class SeaFutureOrder(Frame):
    def __init__(self, information=None):
        Frame.__init__(self)
        self.__obj = dict(
            order = Order(master = self, information = information),
            decrease = DecreaseOrder(master = self, information = information),
            cancel = CancelOrder(master = self, information = information),
        )

    def SetID(self, id):
        for _ in 'order', 'decrease', 'cancel':
            self.__obj[_].SetID(id)

    def SetAccount(self, account):
        for _ in 'order', 'decrease', 'cancel':
            self.__obj[_].SetAccount(account)
