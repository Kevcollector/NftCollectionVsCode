import requests
import json
import time


class ApiAuthor:
    def __init__(self, auther, collection_name):
        self.auther = auther
        self.collection = collection_name
        authers = ("https://proton.api.atomicassets.io/atomicmarket/v1/sales?state=3&account={}&collection_name={}&page=1&limit=100&order=desc&sort=updated".format(self.auther, self.collection))
        authers = requests.get(authers)
        waitUntilReset = int(authers.headers['X-RateLimit-Reset'])
        remainderPings = int(authers.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.authers_ = json.loads(authers.text)
        print("getting Authers sales")

    def update(self, updatetime):
        authers = ("https://proton.api.atomicassets.io/atomicmarket/v1/sales?state=3&account={}&collection_name={}&before={}&page=1&limit=100&order=desc&sort=updated".format(
            self.auther, self.collection, updatetime))
        authers = requests.get(authers)
        waitUntilReset = int(authers.headers['X-RateLimit-Reset'])
        remainderPings = int(authers.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.authers_ = json.loads(authers.text)
        return self.authers_


class ApiResales:
    def __init__(self, auther, collection_name):
        self.auther = auther
        self.collection = collection_name
        resells = ("https://proton.api.atomicassets.io/atomicmarket/v1/sales?state=1%2C3&seller_blacklist={}&buyer_blacklist={}&collection_name={}&page=1&limit=100&order=desc&sort=updated".format(self.auther, self.auther, self.collection))
        resells = requests.get(resells)
        waitUntilReset = int(resells.headers['X-RateLimit-Reset'])
        remainderPings = int(resells.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.resells_ = json.loads(resells.text)
        print("getting Resales")

    def update(self, updatetime):
        resells = ("https://proton.api.atomicassets.io/atomicmarket/v1/sales?state=1%2C3&seller_blacklist={}&buyer_blacklist={}&collection_name={}&before={}&page=1&limit=100&order=desc&sort=updated".format(
            self.auther, self.auther, self.collection, updatetime))
        resells = requests.get(resells)
        waitUntilReset = int(resells.headers['X-RateLimit-Reset'])
        remainderPings = int(resells.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.resells_ = json.loads(resells.text)
        return self.resells_


class ApiOffersAuther:
    pass


class ApiOfferResales:
    pass


class ApiAuctionsAuther:
    pass


class ApiAuctionsResales:
    pass


class ApiTransfersAuther:
    def __init__(self, auther, collection_name):
        self.auther = auther
        self.collection = collection_name
        TranfersAuther = "https://proton.api.atomicassets.io/atomicassets/v1/transfers?collection_name={}&hide_contracts=true&sender={}&page=1&limit=100&order=desc&sort=created".format(
            self.collection, self.auther)
        print(TranfersAuther)
        TranfersAuther = requests.get(TranfersAuther)
        waitUntilReset = int(TranfersAuther.headers['X-RateLimit-Reset'])
        remainderPings = int(TranfersAuther.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.TranfersAuther = json.loads(TranfersAuther.text)
        print("getting your transfers")

    def update(self, time):
        TranfersAuther = "https://proton.api.atomicassets.io/atomicassets/v1/transfers?collection_name={}&hide_contracts=true&sender={}&before={}&page=1&limit=100&order=desc&sort=created".format(
            self.collection, self.auther, time)
        print(TranfersAuther)
        TranfersAuther = requests.get(TranfersAuther)
        waitUntilReset = int(TranfersAuther.headers['X-RateLimit-Reset'])
        remainderPings = int(TranfersAuther.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.TranfersAuther = json.loads(TranfersAuther.text)
        print("getting your transfers (again)")
        return self.TranfersAuther

    def refersh(self, time):
        TranfersAuther = "https://proton.api.atomicassets.io/atomicassets/v1/transfers?collection_name={}&hide_contracts=true&sender={}&after{}&page=1limit=100&order=desc&sort=created".format(
            self.collection, self.auther, time)
        TranfersAuther = requests.get(TranfersAuther)
        waitUntilReset = int(TranfersAuther.headers['X-RateLimit-Reset'])
        remainderPings = int(TranfersAuther.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.TranfersAuther = json.loads(TranfersAuther.text)
        print("getting your transfers (again)")
        return self.TranfersAuther


class ApiTransfersOther:
    def _init__(self, collection_name):
        self.collection = collection_name
        TransfersOther = "https://proton.api.atomicassets.io/atomicassets/v1/transfers?collection_name={}&hide_contracts=true&&page=1&limit=100&order=desc&sort=created".format(
            self.collection)
        TransfersOther = requests.get(TransfersOther)
        waitUntilReset = int(TransfersOther.headers['X-RateLimit-Reset'])
        remainderPings = int(TransfersOther.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.TransfersOther = json.loads(TransfersOther.text)
        print("getting your transfers")

    def update(self, time):
        TransfersOther = "https://proton.api.atomicassets.io/atomicassets/v1/transfers?collection_name={}&hide_contracts=true&before={}&page=1&limit=100&order=desc&sort=created".format(
            self.collection, time)
        TransfersOther = requests.get(TransfersOther)
        waitUntilReset = int(TransfersOther.headers['X-RateLimit-Reset'])
        remainderPings = int(TransfersOther.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.TransfersOther = json.loads(TransfersOther.text)
        print("getting your transfers")
        return TransfersOther

    def refresh(self, time):
        TransfersOther = "https://proton.api.atomicassets.io/atomicassets/v1/transfers?collection_name={}&hide_contracts=true&after={}&page=1&limit=100&order=desc&sort=created".format(
            self.collection, time)
        TransfersOther = requests.get(TransfersOther)
        waitUntilReset = int(TransfersOther.headers['X-RateLimit-Reset'])
        remainderPings = int(TransfersOther.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.TransfersOther = json.loads(TransfersOther.text)
        print("getting your transfers")
        return TransfersOther


class ApiGetBalance:
    pass
