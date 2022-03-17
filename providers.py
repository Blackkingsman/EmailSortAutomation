class Providers:
    def __init__(self, senderemail, count):
        self.senderemail = senderemail
        self.count = count

    def __repr__(self):
        d = dict()
        d["EMAIL"] = self.senderemail
        d["COUNT"] = self.count
        return d

    def getDict(self):
        d = dict()
        d["EMAIL"] = self.senderemail
        d["COUNT"] = self.count
        return d

    def addOne(self):
        self.count += 1

    def getEmail(self):
        return self.senderemail

    def getCount(self):
        return self.count