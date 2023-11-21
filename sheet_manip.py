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

sheet['D4'] = "test"

data = [[],[],[],[]] #list for the extracted data
#extract relevat data from sheet
for row in sheet.iter_rows(min_col=1, max_col=2, values_only=True):
    data[0].append(row[0]) #name
    data[1].append(row[1]) #money 
    print(row)

#convert money to tickets
for val in data[1]:
    tmp = val/20 #convert to tickets
    data[2].append(tmp) #add to data 3 thingy

#calculate total points
ratio = 200
sum = 0
for val in data[2]:
    sum += val

#calculate ratios and place inb slot 3
checksum = 0 #sum to check the averages
for val in data[2]:
    tmp = yeild(val)
    data[3].append(int(round(tmp)))
    checksum += int(round(tmp))
    
if int(checksum) == int(ratio) :
    print("Correct! Total sum equals number of balls")
else : 
    print("Shits Wrong!! is not 200 but is %d" % checksum)
print(data)

#write the tickets to the excel doc
iterator = 1 #iterator running through the for loop
rowstr = "C" #the row the range is to be set at 
for val in data[2]:
    cordtmp = rowstr + str(iterator)
    print("Writing %d to %s" % (val, cordtmp))
    sheet[cordtmp]=val
    iterator += 1

#write the ranges to the excel document
iterator = 1 #iterator running through the for loop
rowstr = "D" #the row the range is to be set at 
for val in data[3]:
    cordtmp = rowstr + str(iterator)
    print("Writing %d to %s" % (val, cordtmp))
    sheet[cordtmp]=val
    iterator += 1
print("Saving as new excel sheet")
workbook.save("liste.xlsx")
print("Finished! sould be written to excel now :)")

