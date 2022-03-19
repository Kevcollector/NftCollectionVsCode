import modules.ApiClass as Api
import modules.SalesClass as Sale
import modules.transfersclass as t
author = "crystalfrogs"
collection_name = "542514111454"

authors = Api.ApiAuthor(author, collection_name)
authors.refresh
sales = Sale.Sales(authors.authors_)
offerResale = Api.ApiOfferResales(author, collection_name)
print(offerResale)
auth = Api.ApiTransfersAuthor(author, collection_name)
others = Api.ApiTransfersAuthor(author, collection_name)
transfers = t.transfers(others.TranfersAuthor)
transfers.printf()
