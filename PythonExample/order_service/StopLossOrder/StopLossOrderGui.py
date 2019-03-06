# 畫視窗用物件
from tkinter import *
from tkinter.ttk import *

# 載入其他物件
from StopLossOrder import FutureStopLossOrder
from StopLossOrder import MovingStopLossOrder
from StopLossOrder import OptionStopLossOrder
from StopLossOrder import FutureOCOOrder
#----------------------------------------------------------------------------------------------------------------------------------------------------

class StopLossOrderGui(Frame):
    def __init__(self, information=None):
        Frame.__init__(self)
        self.__obj = dict(
            future = FutureStopLossOrder.FutureStopLossOrder(master = self, information = information),
            moving = MovingStopLossOrder.MovingStopLossOrder(master = self, information = information),
            option = OptionStopLossOrder.OptionStopLossOrder(master = self, information = information),
            future_oco = FutureOCOOrder.FutureOCOOrder(master = self, information = information),
        )

    def SetID(self, id):
        for _ in 'future', 'moving', 'option', 'future_oco':
            self.__obj[_].SetID(id)

    def SetAccount(self, account):
        for _ in 'future', 'moving', 'option', 'future_oco':
            self.__obj[_].SetAccount(account)

