import pymongo
from pymongo.errors import ConnectionFailure
import traceback
import json


class logging_mongo:
    def connect_mongo(self):
        try:
            # connect to MongoDB, change the << MONGODB URL >> to reflect your own connection string
            return pymongo.MongoClient(
                "mongodb+srv://superadmin:Superadmin_Aspi@aspigrow-01w1b.mongodb.net/appgen-clm?retryWrites=true&w=majority")['appgen-clm']
        except ConnectionFailure as error:
            error_obj = {"isSuccess": False, "message":  str(
                error), "traceback": traceback.format_exc()}
            return json.dumps(error_obj)
