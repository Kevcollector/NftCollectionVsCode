import modules.ApiClass as Api
import modules.SalesClass as Sale
import modules.transfersclass as t
auther = "crystalfrogs"
collection_name = "542514111454"

authers = Api.ApiAuthor(auther, collection_name)
sales = Sale.Sales(authers.authers_)
others = Api.ApiTransfersAuther(auther, collection_name)
transfers = t.transfers(others.TranfersAuther)
transfers.printf()
