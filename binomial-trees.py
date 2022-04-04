import math
import numpy as np
from scipy.optimize import minimize
import xlrd
import xlwt
import pathlib

#defined class node with 2 attributes (price and rate)
class node:
    def __init__(self, price, rate):
        self.price = price
        self.rate = rate

#variable declarations, size = 2 is used to intitalize the base tree
size = 2
branch = []
tree = []
spot_prices = []
imp_vol = []
solved_rates = []


#opens the excel file and reads the sheet
excel_sheet_name = "tree-input.xls" #this file needs to be in the same directory as the python script
path = str(pathlib.Path(__file__).parent.absolute()) + "\\" + excel_sheet_name
book = xlrd.open_workbook(path)
sheet = book.sheet_by_name("Inputs")
data = np.array([[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)])
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('test')
Style = xlwt.easyxf(num_format_str='#,##0.0000000000')

#import the data from the excel sheet here - forward rate, imp vols and zero coupon rates
i = 5
while i < len(data):
    spot_prices.append(float(data[i][2]))
    imp_vol.append(float(data[i][3]))
    i = i + 1
forward_rate = float(data[2][2])

#this is a function that is called by the minimize method it will minimize the difference between
#the spot price at a specific time and the terminal tree price (these should be equal once the tree is sovled)
#the minimize function will call this function while changing the array 'arg' until condition is minimized
#'arg' is an array with only 1 element (the down rate)

def crit_func(arg, fspot_prices, fimp_vol, fcurrent_iter):
    #the tree down rate is changed to equal the 'guess' provided by the minimize function
    tree[1][len(tree)-2].rate = arg[0]
    i = 1
    j = len(tree) - 2 - i

    #iterate through the tree to adjust rates based on the down rate 'guess'
    while j >= 0:
        p = math.exp(2 * fimp_vol[fcurrent_iter] * math.sqrt(.25))
        tree[i][j].rate = tree[i][j+1].rate * p
        j = j - 1
    i = 0

    #iterate through the tree to adjust prices based on new rates
    while i < (len(tree) - 1):
        j = 0
        while j < (len(tree) - i - 1):
            tree[i+1][j].price = (tree[i][j].price + tree[i][j+1].price) * .5 / (1 +  tree[i+1][j].rate/ 4)
            j = j + 1
        i = i + 1

    #define out critical value (what we want to minimize): (spot price) - (terminal tree price)
    crit_value = abs((fspot_prices[fcurrent_iter] - tree[len(tree)-1][0].price))

    #return the critical value to the 'minimize' method for the guess provided by 'arg'
    return crit_value

#this function will create our base tree (with 2 peroids)
def tree_init(fsize, fforward_rate, fimplied_vol):
    i = 0
    fbranch = []
    fprice = 100
    fbranch.append(node(fprice,0))
    fbranch.append(node(fprice,0))
    fbranch.append(node(fprice,0))
    tree.append(fbranch)

    #we will make an estimation of the first up_rate and down_rates (2x and 1/2x)
    fbranch_2 = []
    fbranch_2.append(node(100/(1+fforward_rate/2), fforward_rate*2))
    fbranch_2.append(node(100/(1+fforward_rate/8), fforward_rate/2))
    tree.append(fbranch_2)
    fbranch_3 = []
    fbranch_3.append(node(((fbranch_2[0].price + fbranch_2[1].price)*.5)/(1+fforward_rate/4), fforward_rate))
    tree.append(fbranch_3)


#this will add new branches for each new peroid
#
#Below is a visualization of how the data is structured for the tree using a 2-D array of node objects
#During each iteration an array of price 100 is added to the begining of the array (tree[0][size])
#
#                                                   tree[len-3][0].price
#                           tree[len-2][0].price
#   tree[len-1][0].price                            tree[len-3][1].price
#                           tree[len-2][1].price
#                                                   tree[len-3][2].price
#
#
#                           tree[len-2][0].rate
#   tree[len-1][0].rate
#                           tree[len-2][1].rate
#

def add_branches(current_size):
    new_branch = []
    i = 0
    while i < current_size:
        new_branch.append(node(100,0))
        i = i + 1
    tree.insert(0,new_branch)

#this will run the minimize function to solve the tree of size given by the # of implied vol/spot_prices given in the
#excel spreadsheet
step = 0
while step < len(imp_vol):
    if step == 0:
        tree_init(size, forward_rate, .875)
        X = [tree[1][1].rate]
        y = minimize(crit_func, X, args=(spot_prices,imp_vol, 0), method='nelder-mead', options = {'xatol':.000000000001})
        #print(y)
        step = step + 1
    elif step > 0:
        add_branches(step+3)
        X = [tree[2][step].rate]
        y = minimize(crit_func, X, args=(spot_prices,imp_vol, step), method='nelder-mead', options = {'xatol':.00000000001})
        #print(y)
        step = step + 1

#this will print the tree to an excel spreadsheet with a name "tree-output.xls" in the same folder as the python script
i = len(tree) - 1
l = 1
while i > 0:
    k = 0
    worksheet.col(l).width = 256*20
    while k < i:
        worksheet.write(k*2+l,i,tree[l][k].rate,Style)
        k = k +1
    l = l + 1
    i = i - 1

workbook.save(str(pathlib.Path(__file__).parent.absolute()) + "\\" + "tree-output.xls")
