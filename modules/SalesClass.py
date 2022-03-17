from datetime import datetime

class Sales:
    lists=[]
    def __init__(self,data):
        for x in data['data']:
            self.coin=x["listing_symbol"]
            self.moneyWithoutEdits = x["listing_price"]
            if self.coin=="XUSDC":
                self.moneyUSDC = int(self.moneyWithoutEdits) / 1000000
            else:
                self.moneyUSDC =0
            if self.coin=="XPR":
                self.moneyXPR=int(self.moneyWithoutEdits) / 10000
            else:
                self.moneyXPR =0
            if self.coin=="LOAN":
                self.moneyLOAN = int(self.moneyWithoutEdits) / 10000
            else:
                self.moneyLOAN =0
            if self.coin=="FOOBAR":
                self.moneyFOOBAR = int(self.moneyWithoutEdits) / 1000000
            else:
                self.moneyFOOBAR =0
            self.author=x['collection']["author"]
            self.NFTname = x['assets'][0]['name']
            self.transferTime = datetime.utcfromtimestamp(int(x['assets'][0]['transferred_at_time'])/1000).strftime('%d-%m-%Y %H:%M:%S')
            self.timeMs = x['created_at_time']
            self.number_of_nft = int(x['assets'][0]['template_mint'])
            self.fee = x['collection']['market_fee']
            self.timeSec = int(self.timeMs) / 1000
            self.buyer = x['buyer']
            self.seller = x['seller']
            self.BuyTime = datetime.utcfromtimestamp(self.timeSec).strftime('%d-%m-%Y %H:%M:%S')
            self.lists.append([self.seller, self.moneyUSDC,self.moneyXPR,self.moneyLOAN,self.moneyFOOBAR, self.buyer,self.number_of_nft, self.NFTname,self.BuyTime, self.transferTime])

    def add(self,data):
        def __init__(self):
            super().__init__(data)
    def showSale(self):
        print(f"\n\nNft name {self.NFTname} number {self.number_of_nft}\nprice paid\t{self.coin}\n\t\t\t{self.moneyUSDC}\n\t\t\t{self.moneyXPR}\n\t\t\t{self.moneyLOAN}\n\t\t\t{self.moneyFOOBAR}\n\t\nSeller is {self.seller}\nBuyer is {self.buyer}\nBought at {self.BuyTime}\n\n")
    def showSales(self):
        print(self.lists)
    

