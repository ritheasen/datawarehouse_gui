import pymongo

myclient = pymongo.MongoClient("mongodb://127.0.0.1:27017/")

mydb = myclient["mydatabaseeeee"]
mycol = mydb["customers"]
mydict = { "name": "Jack", "address": "Highway 39" }

# insert
# x = mycol.insert_one(mydict)
x = mycol.find_one()

# update
myquery = { "name" : "Jack" }
newAddress = { "$set" : { "address": "Chennai 123"}}

mycol.update_one(myquery, newAddress)

# delete

myDeleteQuery = { "name" : "John"}
mycol.delete_one(myDeleteQuery)

for x in mycol.find():
  print(x)

# print(x)