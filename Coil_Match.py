#Program by MacGregor Winegard on 7/4/2021
#Armature Coil Matchers
#Takes list of top and bottom armature coils and provides the best pairs
#based on manufacturing data

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import xlsxwriter as xw
    

class coil: #Armature Coil object
    def __init__(self, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S):
        #letter above corresponds to column letter in Excel sheet
        self.bar_num = B #String
        self.Top   = C #Boolean	
        self.CEBB1 = D #All following are doubles
        self.CEBB2 = E
        self.CEBB3 = F
        self.CEBB4 = G
        self.CECA1 = H
        self.CECA2 = I
        self.CECD1 = J 
        self.CECD2 = K
        self.TEBB1 = L
        self.TEBB2 = M
        self.TEBB3 = N
        self.TEBB4 = O
        self.TECA1 = P
        self.TECA2 = Q
        self.TECD1 = R
        self.TECD2 = S
		
        if self.Top == True:
            self.CECA_Avg = (((-1)*self.CECA1) + ((-1)*self.CECA2))*(1/2)
            self.TECA_Avg = (((-1)*self.TECA1) + ((-1)*self.TECA2))*(1/2)
            
        else: #Its a bottom
            self.CECA_Avg = (self.CECA1 + self.CECA2)*(1/2)
            self.TECA_Avg = (self.TECA1 + self.TECA2)*(1/2)
            
            
        self.Delta = abs(self.CECA_Avg - self.TECA_Avg) #may never actully use it but yolo
        self.Sum = self.CECA_Avg + self.TECA_Avg
        self.min_fcn = abs(self.Sum) + abs(self.Delta)

        if self.min_fcn >= 1.6:
            self.color =  "red"
        elif self.min_fcn >=1.2:
            self.color =  "yellow"
        else:
            self.color =  "green"
            
            
            
class Match:
    
    def match(list1, list2, good_first = True):
        if good_first == True:
            list1.sort(key = lambda x: x.min_fcn, reverse = False)
            list2.sort(key = lambda x: x.min_fcn, reverse = False)
        else:
            list1.sort(key = lambda x: x.min_fcn, reverse = True)
            list2.sort(key = lambda x: x.min_fcn, reverse = True)
        
        L1_len = len(list1)
        L2_len = len(list2)
        
        match_list = []
        
        
        while L2_len >0:
            for x in range(L1_len): #go through list 1
                current_Match = 10
                current_Match_loc = -1
                
                current_CE = list1[x].CECA_Avg
                current_TE = list1[x].TECA_Avg
                L2_len = len(list2)
                
                for y in range(L2_len): #go through list 2
                    temp_CE = list2[y].CECA_Avg
                    temp_TE = list2[y].TECA_Avg
                    
                    CE_dif = current_CE - temp_CE
                    TE_dif = current_TE - temp_TE
                    
                    temp_avg_dif = abs((CE_dif + TE_dif)/2)
                    
                    if temp_avg_dif < current_Match:
                        current_Match_loc = y
                        current_Match = temp_avg_dif
                        
                
                temp_list = [list1[x], list2[current_Match_loc], current_Match]
                match_list.append(temp_list)
                
                list2.pop(current_Match_loc)
                L2_len = len(list2)
                
        match_list.sort(key = lambda x: x[2], reverse = False)
        return match_list
        
    def split_full_list(coil_list):
        list_tops = []
        list_bots = []
        
        for x in coil_list: #goes through whole list
            b = x[1]
            d = x[3]
            e = x[4]
            f = x[5]
            g = x[6]
            h = x[7]
            i = x[8]
            j = x[9]
            k = x[10]
            l = x[11] 
            m = x[12] 
            n = x[13] 
            o = x[14] 
            p = x[15] 
            q = x[16] 
            r = x[17] 
            s = x[18]
            
            
            if x[2] == 'T': #extracts top bars
                c = True
                list_tops.append(coil(b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s))   
            elif x[2] == 'B':  #extracts bottom bars
                c = False
                list_bots.append(coil(b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s))
                
        return  [list_tops, list_bots]


class XLWork:
    def XL_to_list():
        root = tk.Tk() #idk what these are 
        root.withdraw()  #https://www.youtube.com/watch?v=H71ts4XxWYU   
        file_path = filedialog.askopenfilename(filetypes = [('Excel Files', '*.xlsx')]) #but basically this opens the file selector  
        df = pd.read_excel(file_path) #https://www.youtube.com/watch?v=S5EVZwXnleM     
        return df.to_numpy() 

    def get_save_path():
        save_loc = filedialog.asksaveasfilename(filetypes = [('Excel Files', 'xlsx.*')]) + '.xlsx'    
        return xw.Workbook(save_loc)  
        
    def write_to_xl(XLPath, pairs_list, trialName):
        
        out_sheet = XLPath.add_worksheet(name = "Compact View -- " + trialName)
        out_sheet.write(0,0, 'Top bar') 
        out_sheet.write(0,1, 'Bottom bar')
        out_sheet.write(0,2, 'Avg Diff Val')   
        for pair in range(len(pairs_list)):
            out_sheet.write(pair+1, 0, pairs_list[pair][0].bar_num)
            out_sheet.write(pair+1, 1, pairs_list[pair][1].bar_num)
            out_sheet.write(pair+1, 2, pairs_list[pair][2])
       
       
        out_sheet = XLPath.add_worksheet(name = "Expanded View -- " + trialName)
        out_sheet.write(0,0, 'Bar #')
        out_sheet.write(0,1, 'T/B')
        out_sheet.write(0,2, 'CE BB1')
        out_sheet.write(0,3, 'CE BB2')
        out_sheet.write(0,4, 'CE BB3')
        out_sheet.write(0,5, 'CE BB4')
        out_sheet.write(0,6, 'CE CA1')
        out_sheet.write(0,7, 'CE CA2')
        out_sheet.write(0,8, 'CE CD1')
        out_sheet.write(0,9, 'CE CD2')
        out_sheet.write(0,10, 'TE BB1')
        out_sheet.write(0,11, 'TE BB2')
        out_sheet.write(0,12, 'TE BB3')
        out_sheet.write(0,13, 'TE BB4')
        out_sheet.write(0,14, 'TE CA1')
        out_sheet.write(0,15, 'TE CA1')
        out_sheet.write(0,16, 'TE CD1')
        out_sheet.write(0,17, 'TE CD2')
        
        location = 2
        for pair in pairs_list:
            out_sheet.write(location, 0, pair[0].bar_num)
            out_sheet.write(location+1, 0, pair[1].bar_num)
            
            
            if (pair[0].Top == True):
                out_sheet.write(location, 1, 'T')
            else: 
                out_sheet.write(location, 1, 'B')
            
            
            if (pair[1].Top == True):
                out_sheet.write(location+1, 1, 'T')
            else: 
                out_sheet.write(location+1, 1, 'B')
       
            
            out_sheet.write(location, 2, pair[0].CEBB1)
            out_sheet.write(location+1, 2, pair[1].CEBB1)
            
            out_sheet.write(location, 3, pair[0].CEBB2)
            out_sheet.write(location+1, 3, pair[1].CEBB2)
            
            out_sheet.write(location, 4, pair[0].CEBB3)
            out_sheet.write(location+1, 4, pair[1].CEBB3)
            
            out_sheet.write(location, 5, pair[0].CEBB4)
            out_sheet.write(location+1, 5, pair[1].CEBB4)
            
            out_sheet.write(location, 6, pair[0].CECA1)
            out_sheet.write(location+1, 6, pair[1].CECA1)
            
            out_sheet.write(location, 7, pair[0].CECA2)
            out_sheet.write(location+1, 7, pair[1].CECA2)
            
            out_sheet.write(location, 8, pair[0].CECD1)
            out_sheet.write(location+1, 8, pair[1].CECD1)
            
            out_sheet.write(location, 9, pair[0].CECD2)
            out_sheet.write(location+1, 9, pair[1].CECD2)
            
            out_sheet.write(location, 10, pair[0].TEBB1)
            out_sheet.write(location+1, 10, pair[1].TEBB1)
            
            out_sheet.write(location, 11, pair[0].TEBB2)
            out_sheet.write(location+1, 11, pair[1].TEBB2)
            
            out_sheet.write(location, 12, pair[0].TEBB3)
            out_sheet.write(location+1, 12, pair[1].TEBB3)
            
            out_sheet.write(location, 13, pair[0].TEBB4)
            out_sheet.write(location+1, 13, pair[1].TEBB4)
            
            out_sheet.write(location, 14, pair[0].TECA1)
            out_sheet.write(location+1, 14, pair[1].TECA1)
            
            out_sheet.write(location, 15, pair[0].TECA2)
            out_sheet.write(location+1, 15, pair[1].TECA2)
            
            out_sheet.write(location, 16, pair[0].TECD1)
            out_sheet.write(location+1, 16, pair[1].TECD1)
            
            out_sheet.write(location, 17, pair[0].TECD2)
            out_sheet.write(location+1, 17, pair[1].TECD2)
            
            out_sheet.write(location+2, 0, 'Avg Diff:')
            out_sheet.write(location+2, 1, pair[2])
            
            location +=4



    
def Main():   
    print("\n================== Matching Armature Coil Pairs! ===========================")
    print("Program written by MacGregor Winegard, son of Edward.\n")
    print("This program is designed to work with the raw excel file from the machine.")
    print("If you have modified the file this program will not do what is intended!")
    input("Press enter to select file: ") #First we need to select the in file       
    
    
    coil_list = XLWork.XL_to_list() #extract data from XL sheet
    [list_tops, list_bots] = Match.split_full_list(coil_list)#split full list into tops and bottoms
    
    good_pair_list = Match.match(list_bots.copy(), list_tops.copy())#make list of pairs
    bad_pair_list = Match.match(list_bots.copy(), list_tops.copy(), good_first = False)#make list of pairs
    
    #Set save location
    print("Now select where you want this to be saved.")
    print("Pleas enter the filename WITHOUT an extension")
    input("Press enter to continue: ")  
    out_wkbk = XLWork.get_save_path() #set path to save to
    XLWork.write_to_xl(out_wkbk, good_pair_list, "Best First")
    XLWork.write_to_xl(out_wkbk, bad_pair_list, "Bad First")
    out_wkbk.close()
    
    print("Done. Have a nice day!")
   
if __name__ == '__main__':
    Main()