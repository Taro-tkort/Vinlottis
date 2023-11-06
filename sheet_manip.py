from openpyxl import load_workbook
import numpy as np

print("Starting sheet manipulator\n")

## FORMULAS ##

#calculate ticket yeild-rate
def yeild(tickets):
    rate = ratio/sum
    stickets = tickets*rate
    print("Calculating rate with: Tickets: %d    Sum: %d    Rate: %d     Total: %d " % (tickets, sum, rate, stickets))
    return stickets



## Code ##

path = "Book.xlsx"
workbook = load_workbook(filename=path)
sheet = workbook.active

data = [[],[],[]] #list for the extracted data
#extract relevat data from sheet
for row in sheet.iter_rows(min_col=1, max_col=2, values_only=True):
    data[0].append(row[0]) #name
    data[1].append(row[1]) #tickets
    print(row)

#calculate total points
ratio = 200
sum = 0
for val in data[1]:
    sum += val

#calculate ratios and place inb slot 3
checksum = 0 #sum to check the averages
for val in data[1]:
    tmp = yeild(val)
    data[2].append(int(round(tmp)))
    checksum += int(round(tmp))
    
if int(checksum) == int(ratio) :
    print("Correct! Total sum equals number of balls")
else : 
    print("Shits Wrong!! is not 200 but is %d" % checksum)
print(data)

#write the ranges to the excel document
iterator = 1 #iterator running through the for loop
rowstr = "C" #the row the range is to be set at 
for val in data[2]:
    cordtmp = rowstr + str(iterator)
    print("Writing %d to %s" % (val, cordtmp))
    sheet['C1']=val
    iterator += 1
print("Finished! sould be written to excel now :)")