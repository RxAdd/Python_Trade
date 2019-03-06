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
        group = LabelFrame(self.__master, text="期貨委託", style="Pink.TLabelframe")
        group.grid(column = 0, row = 0, padx = 10, pady = 10)

        frame = Frame(group, style="Pink.TFrame")
        frame.grid(column = 0, row = 0, padx = 10, pady = 5, sticky = 'ew')

        # 商品代碼
        Label(frame, style="Pink.TLabel", text = "商品代碼").grid(column = 0, row = 0, pady = 3)
            # 輸入框
        txtStockNo = Entry(frame, width = 10)
        txtStockNo.grid(column = 0, row = 1, padx = 10, pady = 3)

        # 買賣別
        Label(frame, style="Pink.TLabel", text = "買賣別").grid(column = 1, row = 0)
            # 輸入框
        boxBuySell = Combobox(frame, width = 5, state='readonly')
        boxBuySell['values'] = Config.BUYSELLSET
        boxBuySell.grid(column = 1, row = 1, padx = 10)

        # 委託條件
        Label(frame, style="Pink.TLabel", text = "委託條件").grid(column = 2, row = 0)
            # 輸入框
        boxPeriod = Combobox(frame, width = 10, state='readonly')
        boxPeriod['values'] = Config.PERIODSET['future']
        boxPeriod.grid(column = 2, row = 1, padx = 10)

        # 當沖與否
        Label(frame, style="Pink.TLabel", text = "當沖與否").grid(column = 3, row = 0)
            # 輸入框
        boxFlag = Combobox(frame, width = 10, state='readonly')
        boxFlag['values'] = Config.FLAGSET['future']
        boxFlag.grid(column = 3, row = 1, padx = 10)

        # 委託價
        Label(frame, style="Pink.TLabel", text = "委託價").grid(column = 4, row = 0)
            # 輸入框
        txtPrice = Entry(frame, width = 12)
        txtPrice.grid(column = 4, row = 1, padx = 10)

        # 委託量
        Label(frame, style="Pink.TLabel", text = "委託量").grid(column = 5, row = 0)
            # 輸入框
        txtQty = Entry(frame, width = 10)
        txtQty.grid(column = 5, row = 1, padx = 10)

        # 倉別
        Label(frame, style="Pink.TLabel", text = "倉別").grid(column = 6, row = 0)
            # 輸入框
        boxNewClose = Combobox(frame, width = 5, state='readonly')
        boxNewClose['values'] = Config.NEWCLOSESET['future']
        boxNewClose.grid(column = 6, row = 1, padx = 10)

        # 盤別
        Label(frame, style="Pink.TLabel", text = "盤別").grid(column = 7, row = 0)
            # 輸入框
        boxReserved = Combobox(frame, width = 10, state='readonly')
        boxReserved['values'] = Config.RESERVEDSET
        boxReserved.grid(column = 7, row = 1, padx = 10)

        # btnSendOrder
        btnSendOrder = Button(frame, style = "Pink.TButton", text = "送出委託")
        btnSendOrder["command"] = self.__btnSendOrder_Click
        btnSendOrder.grid(column = 8, row =  1, padx = 10)

        self.__dOrder['txtStockNo'] = txtStockNo
        self.__dOrder['boxPeriod'] = boxPeriod
        self.__dOrder['boxFlag'] = boxFlag
        self.__dOrder['boxBuySell'] = boxBuySell
        self.__dOrder['txtPrice'] = txtPrice
        self.__dOrder['txtQty'] = txtQty
        self.__dOrder['boxNewClose'] = boxNewClose
        self.__dOrder['boxReserved'] = boxReserved
        
    # 4.下單送出
    # sBuySell, sTradeType, sDayTrade, sNewClose, sReserved
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

            if self.__dOrder['boxPeriod'].get() == "ROD":
                sTradeType = 0
            elif self.__dOrder['boxPeriod'].get() == "IOC":
                sTradeType = 1
            elif self.__dOrder['boxPeriod'].get() == "FOK":
                sTradeType = 2
            
            if self.__dOrder['boxFlag'].get() == "非當沖":
                sDayTrade = 0
            elif self.__dOrder['boxFlag'].get() == "當沖":
                sDayTrade = 1
            
            if self.__dOrder['boxNewClose'].get() == "新倉":
                sNewClose = 0
            elif self.__dOrder['boxNewClose'].get() == "平倉":
                sNewClose = 1
            elif self.__dOrder['boxNewClose'].get() == "自動":
                sNewClose = 2

            if self.__dOrder['boxReserved'].get() == "盤中":
                sReserved = 0
            elif self.__dOrder['boxReserved'].get() == "T盤預約":
                sReserved = 1

            # 建立下單用的參數(FUTUREORDER)物件(下單時要填商品代號,買賣別,委託價,數量等等的一個物件)
            oOrder = sk.FUTUREORDER()
            # 填入完整帳號
            oOrder.bstrFullAccount =  self.__dOrder['boxAccount']
            # 填入期權代號
            oOrder.bstrStockNo = self.__dOrder['txtStockNo'].get()
            # 買賣別
            oOrder.sBuySell = sBuySell
            # ROD、IOC、FOK
            oOrder.sTradeType = sTradeType
            # 非當沖、當沖
            oOrder.sDayTrade = sDayTrade
            # 委託價
            oOrder.bstrPrice = self.__dOrder['txtPrice'].get()
            # 委託數量
            oOrder.nQty = int(self.__dOrder['txtQty'].get())
            # 新倉、平倉、自動
            oOrder.sNewClose = sNewClose
            # 盤中、T盤預約
            oOrder.sReserved = sReserved

            message, m_nCode = skO.SendFutureOrderCLR(self.__dOrder['txtID'], bAsyncOrder, oOrder)
            self.__oMsg.SendReturnMessage("Order", m_nCode, "SendFutureOrderCLR", self.__dOrder['listInformation'])
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
            messagebox.showerror("error！", '請選擇期貨帳號！')            
        else:
            self.__SendOrder_Click(False)

    def __btnSendOrderAsync_Click(self):
        self.__SendOrder_Click(True)

    def __SendOrder_Click(self, bAsyncOrder):
        try:
            message, m_nCode = skO.DecreaseOrderBySeqNo( self.__dOrder['txtID'], bAsyncOrder, self.__dOrder['boxAccount'],\
                self.__dOrder['txtSqlNo'].get(), int(self.__dOrder['txtDecreaseQty'].get()) )
            self.__oMsg.SendReturnMessage("Order", m_nCode, "DecreaseOrderBySeqNo", self.__dOrder['listInformation'])
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

        # 商品代碼、委託序號、委託書號
        row = 0
        self.__radVar = IntVar()

        for _ in 'txtStockNo', 'txtSeqNo', 'txtBookNo':
                # 輸入框
            self.__dOrder[_] = Entry(frame, width = 18, state = 'disabled')
            self.__dOrder[_].grid(column = 1, row = row, padx = 10)

            if _ == 'txtStockNo':
                text = '商品代碼'
            elif _ == 'txtSeqNo':
                text = '委託序號'
            elif _ == 'txtBookNo':
                text = '委託書號'

            rb = Radiobutton(frame, style="Pink.TRadiobutton", text = text, variable = self.__radVar, value = row, command = self.__radCall)
            rb.grid(column = 0, row = row, pady = 3, sticky = 'ew')

            row = row + 1
        self.__dOrder['txtStockNo']['state'] = '!disabled'

        # btnSendOrder
        btnSendOrder = Button(frame, style = "Pink.TButton", text = "送出")
        btnSendOrder["command"] = self.__btnSendOrder_Click
        btnSendOrder.grid(column = 2, row =  2, padx = 50)

    def __radCall(self):
        radSel = self.__radVar.get()
        self.__dOrder['txtStockNo']['state'] = '!disabled' if radSel == 0 else 'disabled'
        self.__dOrder['txtSeqNo']['state'] = '!disabled' if radSel == 1 else 'disabled'
        self.__dOrder['txtBookNo']['state'] = '!disabled' if radSel == 2 else 'disabled'

    # 送出
    def __btnSendOrder_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇期貨帳號！')            
        elif self.__radVar.get() == 0 and self.__dOrder['txtStockNo'].get() == '':
            ans = messagebox.askquestion("提示", '未輸入商品代碼會刪除所有委託單，是否刪單？')
            if ans == 'yes':
                self.__SendOrder_Click(False)
            else:
                return
        elif self.__radVar.get() == 1 and self.__dOrder['txtSeqNo'].get() == '':
            messagebox.showerror("error！", '請輸入欲取消的委託序號！')    
        elif self.__radVar.get() == 2 and self.__dOrder['txtBookNo'].get() == '':
            messagebox.showerror("error！", '請輸入欲取消的委託書號！')            
        else:
            self.__SendOrder_Click(False)

    def __SendOrder_Click(self, bAsyncOrder):
        try:
            if self.__radVar.get() == 0:
                message, m_nCode = skO.CancelOrderByStockNo( self.__dOrder['txtID'], bAsyncOrder, self.__dOrder['boxAccount'],\
                    self.__dOrder['txtStockNo'].get() )
                self.__oMsg.SendReturnMessage("Order", m_nCode, "CancelOrderByStockNo", self.__dOrder['listInformation'])
            elif self.__radVar.get() == 1:
                message, m_nCode = skO.CancelOrderBySeqNo( self.__dOrder['txtID'], bAsyncOrder, self.__dOrder['boxAccount'],\
                    self.__dOrder['txtSeqNo'].get() )
                self.__oMsg.SendReturnMessage("Order", m_nCode, "CancelOrderBySeqNo", self.__dOrder['listInformation'])
            elif self.__radVar.get() == 2:
                message, m_nCode = skO.CancelOrderByBookNo( self.__dOrder['txtID'], bAsyncOrder, self.__dOrder['boxAccount'],\
                    self.__dOrder['txtBookNo'].get() )
                self.__oMsg.SendReturnMessage("Order", m_nCode, "CancelOrderByBookNo", self.__dOrder['listInformation'])
        except Exception as e:
            messagebox.showerror("error！", e)

class FutureOrder(Frame):
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
