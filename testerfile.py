from turtle import clear
import moduals.ApiClass as Api
import moduals.SalesClass as Sale
auther="crystalfrogs"
collection_name="542514111454"
clear()

authers = Api.ApiAuthor(auther, collection_name)
sales = Sale.Sales(authers.authers_)

while len(authers.authers_["data"]) != 0:
    sales.add(authers.authers_)
    authers.update(sales.timeMs)

others=Api.ApiTransfersAuther(auther, collection_name)
others.update("1647029978500")
print(others.TranfersAuther)
