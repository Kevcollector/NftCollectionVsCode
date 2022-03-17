from datetime import datetime


class transfers:
    lists = []

    def __init__(self, data):
        for x in data['data']:
            self.recipient_name = x['recipient_name']
            self.sender_name = x['sender_name']
            self.created_at_time = x["created_at_time"]
            self.memo = x["memo"]
            self.nftNumber = x["assets"][0]["template_mint"]
            self.name = x["assets"][0]["name"]
            self.owner_name = x["assets"][0]["owner"]
            self.FormattedTime = datetime.utcfromtimestamp(
                int(self.created_at_time)/100).strftime('%d-%m-%Y %H:%M:%S')
            self.lists.append([self.name, self.nftNumber, self.sender_name, self.recipient_name,
                               self.memo, self.created_at_time, self.FormattedTime])

    def update(data):
        super().__init__(data)

    def printf(self):
        for selfs in self.lists:
            print(
                f"\n\t\tsender: {selfs[2]}\n\t\trecipient_name: {selfs[3]}\n\t\tmemo: {selfs[4]}\n\t\tname: {selfs[0]}\n\t\tnftNumber: {selfs[1]}\n\t\ttime: {selfs[6]}")

    def print(self):
        for selfs in self.lists:
            print(selfs)

    def show(self):
        print(self.lists)
