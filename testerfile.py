import modules.ApiClass as Api
import modules.SalesClass as Sale
import modules.transfersclass as t
author = "crystalfrogs"
collection_name = "542514111454"

authors = Api.ApiAuthor(author, collection_name)
authors.refresh
sales = Sale.Sales(authors.authors_)
auth = Api.ApiTransfersAuthor(author, collection_name)
others = Api.ApiTransfersOther(collection_name)
transfers = t.transfers(others.TransfersOther)
others = Api.ApiTransfersAuthor(author, collection_name)
transfers = t.transfers(others.TranfersAuthor)
transfers.printf()
