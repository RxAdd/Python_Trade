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
        group = LabelFrame(self.__master, text="證券委託", style="Pink.TLabelframe")
        group.grid(column = 0, row = 0, padx = 10, pady = 10)
        
        frame = Frame(group, style="Pink.TFrame")
        frame.grid(column = 0, row = 0, padx = 10, pady = 5, sticky = 'ew')

        # 商品代碼
        Label(frame, style="Pink.TLabel", text = "商品代碼").grid(column = 0, row = 0, pady = 3)
            # 輸入框
        txtStockNo = Entry(frame, width = 10)
        txtStockNo.grid(column = 0, row = 1, padx = 10, pady = 3)

        # 上市櫃-興櫃
        Label(frame, style="Pink.TLabel", text = "上市櫃-興櫃").grid(column = 1, row = 0)
            #輸入框
        boxPrime = Combobox(frame, width = 10, state='readonly')
        boxPrime['values'] = Config.PRIMESET
        boxPrime.grid(column = 1, row = 1, padx = 10)

        # 買賣別
        Label(frame, style="Pink.TLabel", text = "買賣別").grid(column = 2, row = 0)
            # 輸入框
        boxBuySell = Combobox(frame, width = 10, state='readonly')
        boxBuySell['values'] = Config.BUYSELLSET
        boxBuySell.grid(column = 2, row = 1, padx = 10)

        # 委託條件
        Label(frame, style="Pink.TLabel", text = "委託條件").grid(column = 3, row = 0)
            # 輸入框
        boxPeriod = Combobox(frame, width = 10, state='readonly')
        boxPeriod['values'] = Config.PERIODSET['stock']
        boxPeriod.grid(column = 3, row = 1, padx = 10)

        # 當沖與否
        Label(frame, style="Pink.TLabel", text = "當沖與否").grid(column = 4, row = 0)
            # 輸入框
        boxFlag = Combobox(frame, width = 10, state='readonly')
        boxFlag['values'] = Config.FLAGSET['stock']
        boxFlag.grid(column = 4, row = 1, padx = 10)

        # 委託價
        Label(frame, style="Pink.TLabel", text = "委託價").grid(column = 5, row = 0)
            # 輸入框
        txtPrice = Entry(frame, width = 15)
        txtPrice.grid(column = 5, row = 1, padx = 10)

        # 委託量
        Label(frame, style="Pink.TLabel", text = "委託量").grid(column = 6, row = 0)
            # 輸入框
        txtQty = Entry(frame, width = 10)
        txtQty.grid(column = 6, row = 1, padx = 10)

        # btnSendOrder
        btnSendOrder = Button(frame, style = "Pink.TButton", text = "送出委託")
        btnSendOrder["command"] = self.__btnSendOrder_Click
        btnSendOrder.grid(column = 7, row =  1, padx = 10)
        #SendOrderAsync
        #self.btnSendOrderAsync = Button(self)
        #self.btnSendOrderAsync["text"] = "非同步送單"
        #self.btnSendOrderAsync["command"] = self.__btnSendOrderAsync_Click
        #self.btnSendOrderAsync.grid(column = 7, row = 4)

        self.__dOrder['txtStockNo'] = txtStockNo
        self.__dOrder['boxPrime'] = boxPrime
        self.__dOrder['boxPeriod'] = boxPeriod
        self.__dOrder['boxFlag'] = boxFlag
        self.__dOrder['boxBuySell'] = boxBuySell
        self.__dOrder['txtPrice'] = txtPrice
        self.__dOrder['txtQty'] = txtQty

    # 4.下單送出
    # sPrime, sPeriod, sFlag, sBuySell
    def __btnSendOrder_Click(self):
        if self.__dOrder['boxAccount'] == '':
            messagebox.showerror("error！", '請選擇證券帳號！')            
        else:
            self.__SendOrder_Click(False)

    def __btnSendOrderAsync_Click(self):
        self.__SendOrder_Click(True)

    def __SendOrder_Click(self, bAsyncOrder):
        try:
            if self.__dOrder['boxPrime'].get() == "上市櫃":
                sPrime = 0
            elif self.__dOrder['boxPrime'].get() == "興櫃":
                sPrime = 1
            
            if self.__dOrder['boxPeriod'].get() == "盤中":
                sPeriod = 0
            elif self.__dOrder['boxPeriod'].get() == "盤後":
                sPeriod = 1
            elif self.__dOrder['boxPeriod'].get() == "零股":
                sPeriod = 2
            
            if self.__dOrder['boxFlag'].get() == "現股":
                sFlag = 0
            elif self.__dOrder['boxFlag'].get() == "融資":
                sFlag = 1
            elif self.__dOrder['boxFlag'].get() == "融券":
                sFlag = 2
            elif self.__dOrder['boxFlag'].get() == "無券":
                sFlag = 3
            
            if self.__dOrder['boxBuySell'].get() == "買進":
                sBuySell = 0
            elif self.__dOrder['boxBuySell'].get() == "賣出":
                sBuySell = 1

            # 建立下單用的參數(STOCKORDER)物件(下單時要填股票代號,買賣別,委託價,數量等等的一個物件)
            oOrder = sk.STOCKORDER()
            # 填入完整帳號
            oOrder.bstrFullAccount = self.__dOrder['boxAccount']
            # 填入股票代號
            oOrder.bstrStockNo = self.__dOrder['txtStockNo'].get()
            # 上市、上櫃、興櫃
            oOrder.sPrime = sPrime
            # 盤中、盤後、零股
            oOrder.sPeriod = sPeriod
            # 現股、融資、融券
            oOrder.sFlag = sFlag
            # 買賣別
            oOrder.sBuySell = sBuySell
            # 委託價
            oOrder.bstrPrice = self.__dOrder['txtPrice'].get()
            # 委託數量
            oOrder.nQty = int(self.__dOrder['txtQty'].get())

            message, m_nCode = skO.SendStockOrder(self.__dOrder['txtID'], bAsyncOrder, oOrder)
            self.__oMsg.SendReturnMessage("Order", m_nCode, "SendStockOrder", self.__dOrder['listInformation'])
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
            messagebox.showerror("error！", '請選擇證券帳號！')            
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
            messagebox.showerror("error！", '請選擇證券帳號！')            
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

class StockOrder(Frame):
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
