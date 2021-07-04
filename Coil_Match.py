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
        self.Bar_num = B #String
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
        self.Min_Fcn = abs(self.Sum) + abs(self.Delta)

        if self.Min_Fcn >= 1.6:
            self.color =  "red"
        elif self.Min_Fcn >=1.2:
            self.color =  "yellow"
        else:
            self.color =  "green"
            
            
            
class match:
    def bad_first(list1, list2):
        list1.sort(key = lambda x: x.Min_Fcn, reverse = True)
        list2.sort(key = lambda x: x.Min_Fcn, reverse = True)
        
        L1_len = len(list1)
        L2_len = len(list2)
        
        match_list = []
        
        while L2_len >0:
            for x in range(L1_len): #go through list 1
                current_match = 10
                current_match_loc = -1
                
                current_CE = list1[x].CECA_Avg
                current_TE = list1[x].TECA_Avg
                L2_len = len(list2)
                
                for y in range(L2_len): #go through list 2
                    temp_CE = list2[y].CECA_Avg
                    temp_TE = list2[y].TECA_Avg
                    
                    CE_dif = current_CE - temp_CE
                    TE_dif = current_TE - temp_TE
                    
                    temp_avg_dif = abs((CE_dif + TE_dif)/2)
                    
                    if temp_avg_dif < current_match:
                        current_match_loc = y
                        current_match = temp_avg_dif
                        
                
                temp_list = [list1[x], list2[current_match_loc], current_match]
                
                match_list.append(temp_list)
                
                list2.pop(current_match_loc)
                L2_len = len(list2)
                
        match_list.sort(key = lambda x: x[2], reverse = False)
        print(L2_len)
        return match_list
        
        
    def good_first(list1, list2):
        list1.sort(key = lambda x: x.Min_Fcn, reverse = False)
        list2.sort(key = lambda x: x.Min_Fcn, reverse = False)
        
        L1_len = len(list1)
        L2_len = len(list2)
        
        match_list = []
        
        
        while L2_len >0:
            for x in range(L1_len): #go through list 1
                current_match = 10
                current_match_loc = -1
                
                current_CE = list1[x].CECA_Avg
                current_TE = list1[x].TECA_Avg
                L2_len = len(list2)
                
                for y in range(L2_len): #go through list 2
                    temp_CE = list2[y].CECA_Avg
                    temp_TE = list2[y].TECA_Avg
                    
                    CE_dif = current_CE - temp_CE
                    TE_dif = current_TE - temp_TE
                    
                    temp_avg_dif = abs((CE_dif + TE_dif)/2)
                    
                    if temp_avg_dif < current_match:
                        current_match_loc = y
                        current_match = temp_avg_dif
                        
                
                temp_list = [list1[x], list2[current_match_loc], current_match]
                match_list.append(temp_list)
                
                list2.pop(current_match_loc)
                L2_len = len(list2)
                
        match_list.sort(key = lambda x: x[2], reverse = False)
        return match_list
        
    def splitFullList(coil_list):
        list_tops = []
        list_bots = []
        
        for x in coil_list: #goes through whole list
            if x[2] == 'T': #extracts top bars
                b = x[1]
                c = True
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

                list_tops.append(coil(b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s))
                
            
            elif x[2] == 'B':  #extracts bottom bars
                b = x[1]
                c = False
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
                
                list_bots.append(coil(b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s))
        return  [list_tops, list_bots]


class xlWork:
    def XL2List():
        root = tk.Tk() #idk what these are 
        root.withdraw()  #https://www.youtube.com/watch?v=H71ts4XxWYU   
        file_path = filedialog.askopenfilename(filetypes = [('Excel Files', '*.xlsx')]) #but basically this opens the file selector  
        df = pd.read_excel(file_path) #https://www.youtube.com/watch?v=S5EVZwXnleM     
        return df.to_numpy() 

    def getSavePath():
        save_loc = filedialog.asksaveasfilename(filetypes = [('Excel Files', 'xlsx.*')]) + '.xlsx'    
        return xw.Workbook(save_loc)  
        
    def write2XL(XLPath, pairs_list, trialName):
        
        outSheet = XLPath.add_worksheet(name = "Compact View -- " + trialName)
        outSheet.write(0,0, 'Top bar') 
        outSheet.write(0,1, 'Bottom bar')
        outSheet.write(0,2, 'Avg Diff Val')   
        for item in range(len(pairs_list)):
            outSheet.write(item+1, 0, pairs_list[item][0].Bar_num)
            outSheet.write(item+1, 1, pairs_list[item][1].Bar_num)
            outSheet.write(item+1, 2, pairs_list[item][2])
       
       
        outSheet = XLPath.add_worksheet(name = "Expanded View -- " + trialName)
        outSheet.write(0,0, 'Bar #')
        outSheet.write(0,1, 'T/B')
        outSheet.write(0,2, 'CE BB1')
        outSheet.write(0,3, 'CE BB2')
        outSheet.write(0,4, 'CE BB3')
        outSheet.write(0,5, 'CE BB4')
        outSheet.write(0,6, 'CE CA1')
        outSheet.write(0,7, 'CE CA2')
        outSheet.write(0,8, 'CE CD1')
        outSheet.write(0,9, 'CE CD2')
        outSheet.write(0,10, 'TE BB1')
        outSheet.write(0,11, 'TE BB2')
        outSheet.write(0,12, 'TE BB3')
        outSheet.write(0,13, 'TE BB4')
        outSheet.write(0,14, 'TE CA1')
        outSheet.write(0,15, 'TE CA1')
        outSheet.write(0,16, 'TE CD1')
        outSheet.write(0,17, 'TE CD2')
        
        location = 2
        for item in range(len(pairs_list)):
            outSheet.write(location, 0, pairs_list[item][0].Bar_num)
            outSheet.write(location+1, 0, pairs_list[item][1].Bar_num)
            
            
            if (pairs_list[item][0].Top == True):
                outSheet.write(location, 1, 'T')
            else: 
                outSheet.write(location, 1, 'B')
            
            
            if (pairs_list[item][1].Top == True):
                outSheet.write(location+1, 1, 'T')
            else: 
                outSheet.write(location+1, 1, 'B')
       
            
            outSheet.write(location, 2, pairs_list[item][0].CEBB1)
            outSheet.write(location+1, 2, pairs_list[item][1].CEBB1)
            
            outSheet.write(location, 3, pairs_list[item][0].CEBB2)
            outSheet.write(location+1, 3, pairs_list[item][1].CEBB2)
            
            outSheet.write(location, 4, pairs_list[item][0].CEBB3)
            outSheet.write(location+1, 4, pairs_list[item][1].CEBB3)
            
            outSheet.write(location, 5, pairs_list[item][0].CEBB4)
            outSheet.write(location+1, 5, pairs_list[item][1].CEBB4)
            
            outSheet.write(location, 6, pairs_list[item][0].CECA1)
            outSheet.write(location+1, 6, pairs_list[item][1].CECA1)
            
            outSheet.write(location, 7, pairs_list[item][0].CECA2)
            outSheet.write(location+1, 7, pairs_list[item][1].CECA2)
            
            outSheet.write(location, 8, pairs_list[item][0].CECD1)
            outSheet.write(location+1, 8, pairs_list[item][1].CECD1)
            
            outSheet.write(location, 9, pairs_list[item][0].CECD2)
            outSheet.write(location+1, 9, pairs_list[item][1].CECD2)
            
            outSheet.write(location, 10, pairs_list[item][0].TEBB1)
            outSheet.write(location+1, 10, pairs_list[item][1].TEBB1)
            
            outSheet.write(location, 11, pairs_list[item][0].TEBB2)
            outSheet.write(location+1, 11, pairs_list[item][1].TEBB2)
            
            outSheet.write(location, 12, pairs_list[item][0].TEBB3)
            outSheet.write(location+1, 12, pairs_list[item][1].TEBB3)
            
            outSheet.write(location, 13, pairs_list[item][0].TEBB4)
            outSheet.write(location+1, 13, pairs_list[item][1].TEBB4)
            
            outSheet.write(location, 14, pairs_list[item][0].TECA1)
            outSheet.write(location+1, 14, pairs_list[item][1].TECA1)
            
            outSheet.write(location, 15, pairs_list[item][0].TECA2)
            outSheet.write(location+1, 15, pairs_list[item][1].TECA2)
            
            outSheet.write(location, 16, pairs_list[item][0].TECD1)
            outSheet.write(location+1, 16, pairs_list[item][1].TECD1)
            
            outSheet.write(location, 17, pairs_list[item][0].TECD2)
            outSheet.write(location+1, 17, pairs_list[item][1].TECD2)
            
            outSheet.write(location+2, 0, 'Avg Diff:')
            outSheet.write(location+2, 1, pairs_list[item][2])
            
            location +=4



    
def main():   
    print("\n================== Matching Armature Coil Pairs! ===========================")
    print("Program written by MacGregor Winegard, son of Edward.\n")
    print("This program is designed to work with the raw excel file from the machine.")
    print("If you have modified the file this program will not do what is intended!")
    input("Press enter to select file: ") #First we need to select the in file       
    
    
    coil_list = xlWork.XL2List() #extract data from XL sheet
    [list_tops, list_bots] = match.splitFullList(coil_list)#split full list into tops and bottoms
    
    goodPairList = match.good_first(list_bots.copy(), list_tops.copy())#make list of pairs
    badPairList = match.bad_first(list_bots.copy(), list_tops.copy())#make list of pairs
    
    #Set save location
    print("Now select where you want this to be saved.")
    print("Pleas enter the filename WITHOUT an extension")
    input("Press enter to continue: ")  
    out_wkbk = xlWork.getSavePath() #set path to save to
    xlWork.write2XL(out_wkbk, goodPairList, "Best First")
    xlWork.write2XL(out_wkbk, badPairList, "Bad First")
    out_wkbk.close()
    
    print("Done. Have a nice day!")
   
if __name__ == '__main__':
    main()