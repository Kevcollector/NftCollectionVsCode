import requests
import json
import time


class ApiAuthor:
    def __init__(self, author, collection_name):
        self.author = author
        self.collection = collection_name
        authors = ("https://proton.api.atomicassets.io/atomicmarket/v1/sales?state=3&account={}&collection_name={}&page=1&limit=100&order=desc&sort=updated".format(self.author, self.collection))
        authors = requests.get(authors)
        waitUntilReset = int(authors.headers['X-RateLimit-Reset'])
        remainderPings = int(authors.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.authors_ = json.loads(authors.text)
        print("getting Authors sales")

    def update(self, updatetime):
        authors = ("https://proton.api.atomicassets.io/atomicmarket/v1/sales?state=3&account={}&collection_name={}&before={}&page=1&limit=100&order=desc&sort=updated".format(
            self.author, self.collection, updatetime))
        authors = requests.get(authors)
        waitUntilReset = int(authors.headers['X-RateLimit-Reset'])
        remainderPings = int(authors.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.authors_ = json.loads(authors.text)
        return self.authors_

    def refresh(self, updatetime):
        authors = ("https://proton.api.atomicassets.io/atomicmarket/v1/sales?state=3&account={}&collection_name={}&after={}&page=1&limit=100&order=desc&sort=updated".format(
            self.author, self.collection, updatetime))
        authors = requests.get(authors)
        waitUntilReset = int(authors.headers['X-RateLimit-Reset'])
        remainderPings = int(authors.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.authors_ = json.loads(authors.text)
        return self.authors_


class ApiResales:
    def __init__(self, author, collection_name):
        self.author = author
        self.collection = collection_name
        resells = ("https://proton.api.atomicassets.io/atomicmarket/v1/sales?state=1%2C3&seller_blacklist={}&buyer_blacklist={}&collection_name={}&page=1&limit=100&order=desc&sort=updated".format(self.author, self.author, self.collection))
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
            self.author, self.author, self.collection, updatetime))
        resells = requests.get(resells)
        waitUntilReset = int(resells.headers['X-RateLimit-Reset'])
        remainderPings = int(resells.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.resells_ = json.loads(resells.text)
        return self.resells_

    def refresh(self, updatetime):
        resells = ("https://proton.api.atomicassets.io/atomicmarket/v1/sales?state=1%2C3&seller_blacklist={}&buyer_blacklist={}&collection_name={}&after={}&page=1&limit=100&order=desc&sort=updated".format(
            self.author, self.author, self.collection, updatetime))
        resells = requests.get(resells)
        waitUntilReset = int(resells.headers['X-RateLimit-Reset'])
        remainderPings = int(resells.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.resells_ = json.loads(resells.text)
        return self.resells_


class ApiOffersAuthor:
    pass


class ApiOfferResales:
    pass


class ApiAuctionsAuthor:
    pass


class ApiAuctionsResales:
    pass


class ApiTransfersAuthor:
    def __init__(self, author, collection_name):
        self.author = author
        self.collection = collection_name
        TranfersAuthor = "https://proton.api.atomicassets.io/atomicassets/v1/transfers?collection_name={}&hide_contracts=true&sender={}&page=1&limit=100&order=desc&sort=created".format(
            self.collection, self.author)
        print(TranfersAuthor)
        TranfersAuthor = requests.get(TranfersAuthor)
        waitUntilReset = int(TranfersAuthor.headers['X-RateLimit-Reset'])
        remainderPings = int(TranfersAuthor.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.TranfersAuthor = json.loads(TranfersAuthor.text)
        print("getting your transfers")

    def update(self, time):
        TranfersAuthor = "https://proton.api.atomicassets.io/atomicassets/v1/transfers?collection_name={}&hide_contracts=true&sender={}&before={}&page=1&limit=100&order=desc&sort=created".format(
            self.collection, self.author, time)
        print(TranfersAuthor)
        TranfersAuthor = requests.get(TranfersAuthor)
        waitUntilReset = int(TranfersAuthor.headers['X-RateLimit-Reset'])
        remainderPings = int(TranfersAuthor.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.TranfersAuthor = json.loads(TranfersAuthor.text)
        print("getting your transfers (again)")
        return self.TranfersAuthor

    def refersh(self, time):
        TranfersAuthor = "https://proton.api.atomicassets.io/atomicassets/v1/transfers?collection_name={}&hide_contracts=true&sender={}&after{}&page=1limit=100&order=desc&sort=created".format(
            self.collection, self.author, time)
        TranfersAuthor = requests.get(TranfersAuthor)
        waitUntilReset = int(TranfersAuthor.headers['X-RateLimit-Reset'])
        remainderPings = int(TranfersAuthor.headers['X-RateLimit-Remaining'])
        if remainderPings < 3:
            wait = waitUntilReset-time.time()
            time.sleep(wait)
        self.TranfersAuthor = json.loads(TranfersAuthor.text)
        print("getting your transfers (again)")
        return self.TranfersAuthor


class ApiTransfersOther:
    def _init__(self, collection_name):
        self.collection = collection_name
        TransfersOther = "https://proton.api.atomicassets.io/atomicassets/v1/transfers?collection_name={}&hide_contracts=true&page=1&limit=100&order=desc&sort=created".format(
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
