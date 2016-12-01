import numpy as np
import xlrd
import xlwt

# *************************** Import & Read Data **************************** #
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
# ******************************* Input Data ******************************** #
Input = dic.get('InputData')
Const = dic.get('Constants')
# ************** 5 Modules *************** #
Mod1 = dic.get('NPHX')      # ****NPHX**** #  
Mod2 = dic.get('Jet120W')   # ***Jet120W** #
Mod3 = dic.get('EUHX')      # ****EUHX**** #
Mod4 = dic.get('NHT')       # ****MNHT**** #
Mod5 = dic.get('ITC90W')    # ***ITC90W*** #
#print(Input.item(1,1))
# ******************************* Constants ********************************* #
CT = Const[1,1]; XCFNHT = Const[2,1]; XCFJet120W = Const[3,1]

CP = Const[6,1]; XCEUHX = Const[7,1]; XCITC = Const[8,1]  #XCITC90W
XCFcnJet = Const[9,1]; XCFcnEUHX = Const[10,1]; XCFcnITC = Const[11,1]
XCFcnNPHX =	Const[12,1]; LatJet = Const[13,1]; WidJet = Const[14,1]
LatEUHX = Const[15,1]; WidEUHX = Const[16,1]; LatITC = Const[17,1]
WidITC = Const[18,1]; LatNPHX = Const[19,1]; WidNPHX = Const[20,1]

CE = Const[23,1]; XCPrecip = Const[24,1]; XCTmean = Const[25,1]
XCTPrev = Const[26,1]; ECF = Const[27,1]

Snowto0 = Const[30,1]; WidthS = Const[31,1]; ValueS = Const[32,1];
CS = Const[33,1]; XCFcnTemp = Const[34,1]; XCFcnPrecip = Const[35,1]

CRD = Const[38,1]; XCPrecRD = Const[39,1]; XCJetRD = Const[40,1]	
XCFcnJetRD = Const[41,1]; XCITCRD = Const[42,1]; XCFcnITCRD = Const[43,1]	
WidFcnJetRD = Const[44,1]; LatFcnJetRD = Const[45,1]; WidFcnITC = Const[46,1]	
LatFcnITC = Const[47,1]

WidFcnJetD0 = Const[50,1]; LatFcnJetD0 = Const[51,1]; WidFcnTPD0 = Const[52,1]
LatFcnTPD0 = Const[53,1]; CD0 = Const[54,1]; XCJetD0 = Const[55,1]
XCTPD0 = Const[56,1]; XCFcnJetD0 = Const[57,1]; XCFcnTPD0 = Const[58,1]

WidFcnJetD40 = Const[61,1]; LatFcnJetD40 = Const[62,1]; CD40 = Const[65,1]
WidFcnTPD40 = Const[63,1]; LatFcnTPD40 = Const[64,1]; XCJetD40 = Const[66,1]
XCTPD40 = Const[67,1]; XCFcnJetD40 = Const[68,1]; XCFcnTPD40 = Const[69,1]

# ************************** Input Variables ******************************** #
Tobs = Input[5:17, 4]; NHTmonth = Mod4[4, 2:]; NHTyear = Mod4[5:, 2:]
Jet120Wmonth = Mod2[3, 2:]; Jet120Wyear = Mod2[4:, 2:]

Pobs = Input[5:17, 1]; NPHXmonth = Mod1[2, 2:]; NPHXyear = Mod1[3:, 2:]
EUHXmonth = Mod3[3, 2:]; EUHXyear = Mod3[4:, 2:]; ITCmonth = Mod5[12, 2:]
ITCyear = Mod5[13:, 2:]

Eobs = Input[5:17, 5]; LCTobsm = Tobs[-1]; RCTobsm = Tobs[0:11] #a = a[:index] + a[index+1 :] 
TPrevm = np.append(LCTobsm , RCTobsm)

Sobs = Input[5:17, 6]; RDobs = Input[5:17, 7];
D0obs = Input[5:17, 8]; D40obs = Input[5:17, 9]
#%%
# ************************* Temperature Projection ************************** #
HTemp = np.empty((len(NHTyear),len(NHTyear.T)))
TRawCalcNow = np.empty((np.shape(Tobs)))
for p in range (len(NHTyear.T)):     #12
    for q in range (len(NHTyear)):   #400
        
        TRawCalcNow[p] = (XCFNHT * NHTmonth[p])+(XCFJet120W * Jet120Wmonth[p]) + CT
        
        HTemp[q,p] = CT + Tobs[p] - TRawCalcNow[p] + (XCFNHT*NHTyear[q,p]) + (XCFJet120W*Jet120Wyear[q,p])
        np.set_printoptions(precision=4, suppress=True)
#print(HTemp)
#%%
# ************************ Precipitation Projection ************************* #
PPTCRm = np.power(Pobs,0.33333)
    
HPrecip = np.empty((len(NHTyear),len(NHTyear.T)))
for p in range (len(NHTyear.T)):     #12
    for q in range (len(NHTyear)):   #400
        
        FcnJetmp = np.exp(-0.5*np.power(((LatJet-Jet120Wmonth[p])/WidJet),2))
        FcnEUHXm = np.exp(-0.5*np.power(((LatEUHX-EUHXmonth[p])/WidEUHX),2)) #FcnHigh
        FcnITCmp = np.exp(-0.5*np.power(((LatITC-ITCmonth[p])/WidITC),2))
        FcnNPHXm = np.exp(-0.5*np.power(((LatNPHX-NPHXmonth[p])/WidNPHX),2)) #Other
        PRawm = np.power((CP + XCEUHX * EUHXmonth[p] + XCITC * ITCmonth[p] + XCFcnJet * FcnJetmp + XCFcnEUHX * FcnEUHXm + XCFcnITC * FcnITCmp + XCFcnNPHX * FcnNPHXm),3)
        #print(PRawm)
        FcnJetyp = np.exp(-0.5*np.power(((LatJet-Jet120Wyear[q,p])/WidJet),2))
        FcnEUHXy = np.exp(-0.5*np.power(((LatEUHX-EUHXyear[q,p])/WidEUHX),2)) #FcnHigh
        FcnITCyp = np.exp(-0.5*np.power(((LatITC-ITCyear[q,p])/WidITC),2))
        FcnNPHXy = np.exp(-0.5*np.power(((LatNPHX-NPHXyear[q,p])/WidNPHX),2)) #Other
    
        HPrecip[q,p]=(np.power((CP + (XCEUHX * EUHXyear[q,p]) + (XCITC * ITCyear[q,p]) + (XCFcnJet * FcnJetyp) + (XCFcnEUHX * FcnEUHXy) + (XCFcnITC * FcnITCyp) + (XCFcnNPHX * FcnNPHXy)),3)) * (Pobs[p] / PRawm)
        np.set_printoptions(precision=2, suppress=True)
#print(HPrecip)
#%%
# ************************* Evaporation Projection ************************** #
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
#print (HEvap)
#%%
# **************************** Snow Projection ****************************** #
FcnPrecipm = np.exp(-0.5*np.power([(ValueS - X)/WidthS  for X in Pobs],2))
FcnPrecipy = np.exp(-0.5*np.power([(ValueS - Y)/WidthS  for Y in HPrecip],2))
RatioStoP = Sobs/Pobs

FcnTempm = np.empty((np.shape(Tobs)))
FcnTempy = np.empty((np.shape(HTemp)))
HSnow = np.empty((np.shape(HPrecip)))

for p in range (len(HPrecip.T)):     #12
    for q in range (len(HPrecip)):   #400
        
        if Tobs[p] < Snowto0:
            FcnTempm[p] = -1 * (Tobs[p]-Snowto0)
        else: FcnTempm[p] = 0
        
        SrawCalNow = CS + (XCFcnTemp * FcnTempm) + (XCFcnPrecip * FcnPrecipm)
        
        if HTemp[q,p] < Snowto0:
            FcnTempy[q,p] = -1 * (HTemp[q,p]-Snowto0)
            HSnow[q,p] = (CS + (XCFcnTemp * FcnTempy[q,p]) + (XCFcnPrecip * FcnPrecipy[q,p])) * (Sobs[p]/SrawCalNow[p])
        else: 
            HSnow [q,p] = 0 
            FcnTempy[q,p] = 0
#print (HSnow)
#%%
# ******************** Number of Rain Days Projection ************************#
PPTOverRD = Pobs/RDobs
FcnJetmRD = np.exp(-0.5*np.power([(LatFcnJetRD - X)/WidFcnJetRD for X in Jet120Wmonth],2))      
FcnITCmRD = np.exp(-0.5*np.power([(LatFcnITC - Y)/WidFcnITC for Y in ITCmonth],2))
RDrawCalNow = CRD + (XCPrecRD * Pobs) + (XCJetRD * Jet120Wmonth) + (XCFcnJetRD * FcnJetmRD) + (XCITCRD * ITCmonth) +(XCFcnITCRD * FcnITCmRD)

FcnJetyRD = np.exp(-0.5*np.power([(LatFcnJetRD - X)/WidFcnJetRD for X in Jet120Wyear],2))      
FcnITCyRD = np.exp(-0.5*np.power([(LatFcnITC - Y)/WidFcnITC for Y in ITCyear],2))

HRainDay = np.empty((np.shape(HPrecip)))
for p in range (len(HPrecip.T)):     
    for q in range (len(HPrecip)):
        HRainDay[q,p] = (CRD + (XCPrecRD * HPrecip[q,p]) + (XCJetRD * Jet120Wyear[q,p]) + (XCFcnJetRD * FcnJetyRD[q,p]) + (XCITCRD * ITCyear[q,p]) +(XCFcnITCRD * FcnITCyRD[q,p])) * (PPTOverRD[p]/RDrawCalNow[p])
        #FcnJetyRD = np.exp(-0.5*np.power(((LatFcnJet - Jet120Wyear[q,p])/WidFcnJet),2))
        #FcnITCyRD = np.exp(-0.5*np.power(((LatFcnITC - ITCyear[q,p])/WidFcnITC),2))
#print (HRainDay)
#%%
# ************ Number of Days with Temp below 0C Projection ******************#
FcnJetmD0 = np.exp(-0.5*np.power([(LatFcnJetD0 - X)/WidFcnJetD0 for X in Jet120Wmonth],2))      
FcnTPmD0 = np.exp(-0.5*np.power([(LatFcnTPD0 - Y)/WidFcnTPD0 for Y in TRawCalcNow],2))

D0rawCalNow = CD0 + (XCTPD0 * TRawCalcNow) + (XCJetD0 * Jet120Wmonth) + (XCFcnJetD0 * FcnJetmD0) + (XCFcnTPD0 * FcnTPmD0)

FcnJetyD0 = np.exp(-0.5*np.power([(LatFcnJetD0 - X)/WidFcnJetD0 for X in Jet120Wyear],2))      
FcnTPyD0 = np.exp(-0.5*np.power([(LatFcnTPD0 - Y)/WidFcnTPD0 for Y in HTemp],2))

HDay0 = np.empty((np.shape(HTemp)))
for p in range (len(HTemp.T)):     
    for q in range (len(HTemp)):
        HDay0[q,p] = (CD0 + (XCTPD0 * HTemp[q,p]) + (XCJetD0 * Jet120Wyear[q,p]) + (XCFcnJetD0 * FcnJetyD0[q,p]) + (XCFcnTPD0 * FcnTPyD0[q,p])) * (D0obs[p]/D0rawCalNow[p])
#print (HDay0)
#%%
# ********** Number of Days with Temp Higher than 40C Projection *************#
FcnJetmD40 = np.exp(-0.5*np.power([(LatFcnJetD40 - X)/WidFcnJetD40 for X in Jet120Wmonth],2))      
FcnTPmD40 = np.exp(-0.5*np.power([(LatFcnTPD40 - Y)/WidFcnTPD40 for Y in TRawCalcNow],2))

D40rawCalNow = CD40 + (XCTPD40 * TRawCalcNow) + (XCJetD40 * Jet120Wmonth) + (XCFcnJetD40 * FcnJetmD40) + (XCFcnTPD40 * FcnTPmD40)

FcnJetyD40 = np.exp(-0.5*np.power([(LatFcnJetD40 - X)/WidFcnJetD40 for X in Jet120Wyear],2))      
FcnTPyD40 = np.exp(-0.5*np.power([(LatFcnTPD40 - Y)/WidFcnTPD40 for Y in HTemp],2))

HDay40 = np.empty((np.shape(HTemp)))
for p in range (len(HTemp.T)):     
    for q in range (len(HTemp)):
        HDay40[q,p] = (CD40 + (XCTPD40 * HTemp[q,p]) + (XCJetD40 * Jet120Wyear[q,p]) + (XCFcnJetD40 * FcnJetyD40[q,p]) + (XCFcnTPD40 * FcnTPyD40[q,p])) * (D40obs[p]/D40rawCalNow[p])
#print (HDay40[0:400, 4:9])
#%%
# ********************** Export Output in an Excel file **********************#
Name = ['HTemp', 'HPrecip', 'HEvap', 'HSnow', 'HRainDay', 'HDayL0', 'HDayH40']
Variable = [HTemp, HPrecip, HEvap, HSnow, HRainDay, HDay0, HDay40]
SizeVar = np.array(np.shape(Variable))
Output = xlwt.Workbook()
def Write2XLS(Sheetname,ProjectedData):
    Sheet = Output.add_sheet(Sheetname)
    Month = np.array(["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"])
    Sheet.write(0, 0, "100PY/Month")
    for p in range (SizeVar[2]): 
        Sheet.write(0, p+1, Month[p])
    for q in range (SizeVar[1]):
        PYear = -100*q
        Sheet.write(q+1, 0, PYear) #ColPYear
    NumFormat = xlwt.easyxf(num_format_str='0.00')    
    for p in range (SizeVar[2]):     
        for q in range (SizeVar[1]):
            Sheet.write(q+1, p+1, ProjectedData[q,p], NumFormat)
    Output.save("Output.xls")
    return
for u in range (SizeVar[0]):
    Write2XLS(Name[u],Variable[u])
#%%
# ************************* Seasonal Projection ******************************#
SeasonalTemp = np.empty((len(HTemp),4))
for p in range (4):
    for q in range (len(HTemp)):
        SeasonalTemp[q,p] = HTemp[q,3*p:3*(p+1)].mean()
#print (SeasonalTemp)
SeasonalSVariable = [HPrecip, HEvap, HSnow, HRainDay, HDay0, HDay40]
SizeSVar = np.array(np.shape(SeasonalSVariable))
SeasonalSumVar = np.empty((len(HPrecip), 4, SizeSVar[0]))
for v in range (SizeSVar[0]):
    for p in range (4):
        for q in range (len(HPrecip)):
            VariableVth = SeasonalSVariable [v]
            SeasonalSumVar[q,p,v] = VariableVth[q, 3*p:3*(p+1)].sum()
SizeSSLSumVar = np.array(np.shape(SeasonalSumVar)) 
OutputSS = xlwt.Workbook()
Sheet2 = OutputSS.add_sheet('Sheetname2')
Season = np.array(["Winter", "Spring", "Summer", "Autumn"])
    Sheet.write(0, 0, "100PY/Month")
NumFormat = xlwt.easyxf(num_format_str='0.00')    
for p in range (SizeSSLSumVar[1]):
    for q in range (SizeSSLSumVar[0]):
        Sheet2.write(q,p,SeasonalSumVar[q,p,5])
OutputSS.save("OutputSS.xls") 


SSLClimVar = np.dstack([SeasonalSumVar,SeasonalTemp])
SizeSSLClimVar = np.array(np.shape(SSLClimVar))





























