import csv

reader = csv.DictReader(open("file2.csv"))
for raw in reader:
    done = "'Lynxy #{}':{},".format(raw["ID"], raw["Rarity Score"])
    print(done,end='')
