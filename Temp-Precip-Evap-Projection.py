import numpy as np
import xlrd

# *********** Import & Read Data ********** #
Data1 = xlrd.open_workbook('Data.xlsx')
nsheet = Data1.nsheets                
sheetnames = Data1.sheet_names()

def f(x,y,z):
        sheetz = Data1.sheet_by_index(z)
        Values = np.array(sheetz.cell(x,y).value)
        return Values
        
dic = {}
for k in range (0,nsheet):
    sheetnamez=sheetnames[k]
    sheetz = Data1.sheet_by_index(k)
    #fz = f(k)
    R = np.array(sheetz.nrows)
    C = np.array(sheetz.ncols)
    
    M = np.empty((R,C))
    dt = np.empty((R,C))
    
    for i in range(R): 
        for j in range(C): 
            #print (i,j)
            h = f(i,j,k)
            #print (z)
            dt = h.dtype
            #print (dt)
            if dt!='float64':
                M[i,j]=np.array(-999999)
            else: M[i,j]=h
            np.set_printoptions(precision=3, suppress=True)
            dic[sheetnamez]=M
#print (M)
#print (dic) 
#%%
# ************** Input Data ************** #
Input = dic.get('InputData')
# ************** 5 Modules *************** #
Mod1 = dic.get('NPHX')      # ****NPHX**** #  
Mod2 = dic.get('Jet120W')   # ***Jet120W** #
Mod3 = dic.get('EUHX')      # ****EUHX**** #
Mod4 = dic.get('NHT')       # ****MNHT**** #
Mod5 = dic.get('ITC90W')    # ***ITC90W*** #
#print(Input.item(1,1))
# ************** Constants *************** #
Const = dic.get('Constants')

# ************* Constants **************** #
CT = Const[1,1]
XCFNHT = Const[2,1]
XCFJet120W = Const[3,1]

CP = Const[6,1]
XCEUHX = Const[7,1]	
XCITC = Const[8,1]   #XCITC90W
XCFcnJet = Const[9,1]
XCFcnEUHX = Const[10,1]
XCFcnITC = Const[11,1]
XCFcnNPHX =	Const[12,1]
LatJet = Const[13,1]
WidJet = Const[14,1]
LatEUHX = Const[15,1]
WidEUHX = Const[16,1]
LatITC = Const[17,1]
WidITC = Const[18,1]
LatNPHX = Const[19,1]
WidNPHX = Const[20,1]

CE = Const[23,1]
XCPrecip = Const[24,1]
XCTmean = Const[25,1]
XCTPrev = Const[26,1]
ECF = Const[27,1]

# ***************** Input Variables ***************** #
Tobs = Input[5:17 ,4]
NHTmonth = Mod4[4, 2:] 
NHTyear = Mod4[5:, 2:]
Jet120Wmonth = Mod2[3, 2:]
Jet120Wyear = Mod2[4:, 2:]

Pobs = Input[5:17 ,1]
NPHXmonth = Mod1[2, 2:]
NPHXyear = Mod1[3:, 2:]
EUHXmonth = Mod3[3, 2:]
EUHXyear = Mod3[4:, 2:]
ITCmonth = Mod5[12, 2:]
ITCyear = Mod5[13:, 2:]

Eobs = Input[5:17 ,5]
LCTobsm = Tobs[-1]
RCTobsm = Tobs[0:11] #a = a[:index] + a[index+1 :] 
TPrevm = np.append(LCTobsm , RCTobsm)
#print(TPrev)
#%%
# **************** Temperature Projection **************** #
HTemp = np.empty((len(NHTyear),len(NHTyear.T)))
for p in range (len(NHTyear.T)):     #12
    for q in range (len(NHTyear)):   #400
        TRawCalcNow = (XCFNHT * NHTmonth[p])+(XCFJet120W * Jet120Wmonth[p]) + CT
        HTemp[q,p] = CT + Tobs[p] - TRawCalcNow + (XCFNHT*NHTyear[q,p]) + (XCFJet120W*Jet120Wyear[q,p])
        np.set_printoptions(precision=2, suppress=True)
#print(HTemp)
#%%
# *************** Precipitation Projection *************** #
PPTCRm = np.power(Pobs,0.33333)
    
HPrecip = np.empty((len(NHTyear),len(NHTyear.T)))

for p in range (len(NHTyear.T)):     #12
    for q in range (len(NHTyear)):   #400
        
        FcnJetm = np.exp(-0.5*np.power(((LatJet-Jet120Wmonth[p])/WidJet),2))
        FcnEUHXm = np.exp(-0.5*np.power(((LatEUHX-EUHXmonth[p])/WidEUHX),2)) #FcnHigh
        FcnITCm = np.exp(-0.5*np.power(((LatITC-ITCmonth[p])/WidITC),2))
        FcnNPHXm = np.exp(-0.5*np.power(((LatNPHX-NPHXmonth[p])/WidNPHX),2)) #Other
        PRawm = np.power((CP + XCEUHX * EUHXmonth[p] + XCITC * ITCmonth[p] + XCFcnJet * FcnJetm + XCFcnEUHX * FcnEUHXm + XCFcnITC * FcnITCm + XCFcnNPHX * FcnNPHXm),3)
        #print(PRawm)
        FcnJety = np.exp(-0.5*np.power(((LatJet-Jet120Wyear[q,p])/WidJet),2))
        FcnEUHXy = np.exp(-0.5*np.power(((LatEUHX-EUHXyear[q,p])/WidEUHX),2)) #FcnHigh
        FcnITCy = np.exp(-0.5*np.power(((LatITC-ITCyear[q,p])/WidITC),2))
        FcnNPHXy = np.exp(-0.5*np.power(((LatNPHX-NPHXyear[q,p])/WidNPHX),2)) #Other
    
        HPrecip[q,p]=(np.power((CP + (XCEUHX * EUHXyear[q,p]) + (XCITC * ITCyear[q,p]) + (XCFcnJet * FcnJety) + (XCFcnEUHX * FcnEUHXy) + (XCFcnITC * FcnITCy) + (XCFcnNPHX * FcnNPHXy)),3)) * (Pobs[p] / PRawm)
        np.set_printoptions(precision=2, suppress=True)
#print(HPrecip)
#%%
# **************** Evaporation Projection **************** #
LCHTempy = HTemp[:,11]
RCHTempy = HTemp[:,0:11]
TPrevy = np.c_[LCHTempy,RCHTempy]
#print (TPrevy)

HEvap = np.empty((len(HPrecip),len(HPrecip.T)))
for p in range (len(HPrecip.T)):     #12
    for q in range (len(HPrecip)):   #400
        ERawm = np.power((CE + XCPrecip * Pobs[p] + XCTmean * Tobs[p] + XCTPrev * TPrevm[p]),3)
        HEvap[q,p]=ECF * (np.power((CE + (XCPrecip * HPrecip[q,p]) + (XCTmean * HTemp[q,p]) + (XCTPrev * TPrevy[q,p])),3)) * (Eobs[p] / ERawm)
        np.set_printoptions(precision=2, suppress=True)
print (HEvap)





















