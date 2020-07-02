import xlrd
file = xlrd.open_workbook('EAMCET2018LASTRANK-converted.xlsx')
sheet = file.sheet_by_index(0)
li = []
sheet.cell_value(0,0)
for i in range(1,sheet.nrows):
    values = sheet.row_values(i)
    li.append([values[0],values[2],values[9],values[10],values[11],values[23],values[1]])
print(li[0])
print(len(li))
l = len(li[0])
types = [type(li[0][0]),type(li[0][1]),type(li[0][2]),type(li[0][3]),type(li[0][4])]
for i in range(len(li)):
    if(li[i][5]==''):
        li[i][5] = 1000000.0
    if(li[i][4]==''):
        li[i][4] = 1000000.0
    if(types[0] != type(li[i][0]) or types[1] != type(li[i][1]) or types[2] != type(li[i][2]) or types[3] != type(li[i][3]) or types[4] != type(li[i][4])):
        print(' ')    
li = sorted(li,key=lambda li: li[4])
cse = []
ece =[]
eee = []
mec = []
for i in range(len(li)):
    if li[i][3] == 'CSE':
        cse.append([li[i][1],li[i][4],li[i][5],li[i][-1]])
    if li[i][3] == 'ECE':
        ece.append([li[i][1],li[i][4],li[i][5],li[i][-1]])
    if li[i][3] == 'EEE':
        eee.append([li[i][1],li[i][4],li[i][5],li[i][-1]])
    if li[i][3] == 'MEC':
        mec.append([li[i][1],li[i][4],li[i][5],li[i][-1]])
cse.sort(key = lambda li:li[1])
ece.sort(key = lambda li:li[1])
eee.sort(key = lambda li:li[1])
mec.sort(key= lambda li: li[1])
rank = []
for i in range(len(cse)):
    rank.append([cse[i][0],[i],cse[i][-1]])

for i in range(len(ece)):
    for clg in range(len(rank)):
        if rank[clg][0]==ece[i][0]:
               rank[clg][1].append(i)
               break
for i in range(len(eee)):
    for clg in range(len(rank)):
        if rank[clg][0]==eee[i][0]:
               rank[clg][1].append(i)
               break
for i in range(len(mec)):
    for clg in range(len(rank)):
        if rank[clg][0]==mec[i][0]:
               rank[clg][1].append(i)
               break
for i in range(len(rank)):
    rank[i].append(sum(rank[i][1])/len(rank[i][1]))
rank.sort(key = lambda a : sum(a[1])/len(a[1]))
for i,rank in zip(range(len(rank)),rank):
    print(i,rank)
