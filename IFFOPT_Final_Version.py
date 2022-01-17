#!/usr/bin/env python
# coding: utf-8

# In[130]:


import win32com.client
import os
import subprocess
import datetime
import time
import pandas as pd

import pyomo.environ as pe
import pyomo.opt as po
import numpy as np
import random
import win32com.client as win32

from collections import defaultdict

import logging
logging.getLogger('pyomo.core').setLevel(logging.ERROR)
USERNAME='DELIOFR' #SAP user
PASSWORD='Branchbound8' #SAP password
path_exe = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
file_generation_path='C:\\Users\\deliofr\\OneDrive - Ecolab\\Documents\\Datasets' 
#folder where the data is to be exported from SAP 
VL060_file_name='TestingDeliveries'+str(datetime.date.today().strftime('%d.%m.%Y'))+'.txt'
ME2N_file_name="TestingDates"+str(datetime.date.today().strftime('%d.%m.%Y'))+".txt"
file_generation_path='C:\\Users\\frand\\Documents'
#VL060_file_name='TestingDeliveries10.12.2021.txt'
#ME2N_file_name='TestingDates10.12.2021.txt'


# In[17]:


def SapExtractionStage(): #process detailed on documentation. Data extraction from SAP visiting 2 transactions.
    
        
        process=subprocess.Popen(path_exe)
        time.sleep(5)
        
        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        
        connection = application.OpenConnection("101. PEE-EBS-ECC Production Operations", True)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return
        
        #login
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = USERNAME
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = PASSWORD
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 9
        session.findById("wnd[0]").sendVKey(0)
        
        #extraction
        session.findById("wnd[0]/tbar[0]/okcd").text = "VL06O"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/btnBUTTON6").press()
        
        session.findById("wnd[0]/usr/ctxtIF_VSTEL-LOW").text = "BE41"
        session.findById("wnd[0]/usr/ctxtIT_ERDAT-LOW").text = datetime.date.today().strftime('%d.%m.%Y')#datetime.date.today().strftime('%d.%m.%Y')
        #session.findById("wnd[0]/usr/txtIT_ERNAM-LOW").text  = "DIDDEKA"
        session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").text = ""
        session.findById("wnd[0]/usr/ctxtIT_LFART-LOW").text = "ZNL"
        session.findById("wnd[0]/usr/ctxtIT_LFART-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtIT_LFART-LOW").caretPosition = 3
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        session.findById("wnd[0]/tbar[1]/btn[18]").press()
        session.findById("wnd[0]").sendVKey(33)
        time.sleep(1)
        session.findById("wnd[1]/tbar[0]/btn[71]").press()
        session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = False
        session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "IFFOPT"
        session.findById("wnd[2]/usr/chkSCAN_STRING-START").setFocus()
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        time.sleep(2)
        session.findById("wnd[3]/usr/sub/1[0,0]/sub/1/3[0,2]/lbl[2,2]").setFocus()
        session.findById("wnd[3]/usr/sub/1[0,0]/sub/1/3[0,2]/lbl[2,2]").caretPosition = 6
        session.findById("wnd[3]").sendVKey(2)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        
        session.findById("wnd[0]").sendVKey(45)
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = file_generation_path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = VL060_file_name
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 21
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(2)
        try:
                
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
                
        except:
                
            print('New file.')
            
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(5)
        
        deliveries=pd.read_csv(r'%s' % file_generation_path+'\\'+VL060_file_name,sep="\t", thousands=r'.').iloc[:,1:] #import feasible STOs data
        deliveries=deliveries.loc[deliveries['Delivery quantity']>0,:] #check for null deliveries?
        deliveries.loc[:,'Purch.Doc.'].to_clipboard(excel=True,index=False,header=False) #get ids for STOs
        
        right_export=False
        i=0
        
        while (right_export==False) & (i<6):
            
            session.findById("wnd[0]/tbar[0]/okcd").text = "ME2N"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/btn%_EN_EBELN_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[24]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[0]").sendVKey(23)

            session.findById("wnd[0]/tbar[1]/btn[33]").press()
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 55+i
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = str(55+i)

            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

            session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = file_generation_path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ME2N_file_name
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            try:
                
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                
            except:
                
                print('New file.')
            
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
            #session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            dates=pd.read_csv(r'%s' % file_generation_path+'\\'+ME2N_file_name,sep='\t',skiprows=3).iloc[:,1:]
            
            try:
                
                if np.sum(dates.columns==['Material', 'Purch.Doc.', 'Del. Date', 'Short Text'])==4:

                    right_export=True
                    print('Right dates imported.')

            except:

                print('Wrong dates imported initially.')
            
            i+=1
            
        if i>=3:
            
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'franciscodelioperego@gmail.com'
            mail.Subject = 'ME2N SAP Bot disadjusted.'
            mail.Body = str(i)+'positions off. Need to add.'

            # To attach a file to the email (optional):
            mail.Send()
            print('Urgent ME2N custom layout search')
        
        elif i==5:
            
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'franciscodelioperego@gmail.com'
            mail.Subject = 'ME2N extraction error.'
            mail.Body = 'What they said.'

            # To attach a file to the email (optional):
            mail.Send()
            print('Urgent ME2N custom layout search')
        
        #process.terminate()
        
        
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        
        return

SapExtractionStage()


# In[92]:


smat=pd.read_excel('Segregation tool open file - Excel Fran FR.xlsx', sheet_name='New Matrix',nrows=37,usecols='A:AL')
#Importing conflicts matrix. Cartesian product of segregation groups that maps to 1 were there a conflict, else to 0.
smat=smat.rename(columns={'Unnamed: 0':'Class'})
smat=smat.set_index('Class')
smat=smat+smat.transpose()
#We make the matrix symmetric. Working with a symmetric matrix is more efficient because of technicalities relating to software use.

materials=pd.read_excel('Segregation tool open file - Excel Fran FR.xlsx', sheet_name='Data sheet SAP') #import materials data
materials=materials.iloc[1:,0:]
materials.loc[pd.isna(materials["Segregation group"]),"Segregation group"]="None"
materials.loc[materials["Segregation group"]=='8AC',"Segregation group"]="8Ac"
materials.loc[pd.isna(materials["Pallet type Essers "]),
              "Pallet type Essers "]=materials.loc[pd.isna(materials["Pallet type Essers "]),"Pallet type SAP"]
materials.loc[materials.loc[:,"Material"]==3112260,"Gross Weight"]=12.51


# In[93]:


deliveries=pd.read_csv(r'%s' % file_generation_path+'\\'+VL060_file_name,sep="\t", encoding='latin', thousands=r',').iloc[:,1:9] 
#importing STO lines for which there's supply availability alognside relevant information e.g SKU, units requested, STO mother.
deliveries2=pd.read_csv(r'%s' % file_generation_path+'\\'+VL060_file_name, sep="\t", encoding='latin').iloc[:,9:11] #import feasible STOs data
#importing relevant information to Lead times apart because the format requires it
deliveries=pd.concat([deliveries, deliveries2], axis=1)
deliveries['Lead Times']=(pd.to_datetime(deliveries['Deliv.date'], format='%d.%m.%Y')-pd.to_datetime(deliveries['Plan GI Dt'], format='%d.%m.%Y')).astype('timedelta64[D]')
#Lead times calculation
dates=pd.read_csv(r'%s' % file_generation_path+'\\'+ME2N_file_name,sep='\t',skiprows=3).iloc[:,1:]
#importing mother STOs to get appropiate delivery dates

from datetime import date
dates['Del. Date'] = pd.to_datetime(dates['Del. Date'], format='%d.%m.%Y')
dates['Del. Date']=dates['Del. Date'].apply(lambda x: (np.datetime64(date.today())-x).days)
#due delivery dates expressed in terms of -(number of days until due delivery date from today)
dates=dates.drop(columns="Short Text")
deliveries


# In[94]:


new=pd.merge(deliveries,materials[["Material","Pallet type Essers ","Breedte Essers ","Gross Weight"
                ,"Net Weight","GR slips Essers","Segregation group"]], how='inner', on="Material")
new["NPallets"]=new['Delivery quantity']/new["GR slips Essers"]
pd.pivot_table(new, values='NPallets', index='Name of sold-to party', columns='Pallet type Essers ', aggfunc='sum').fillna(0)
#pallets worth of volume per type and location


# In[95]:


deliveries.loc[deliveries.loc[:,"DPrio"]==55,:].iloc[:,3].value_counts() #high priority STO lines by location


# ## Optimization Model:
# \
#     \begin{equation*}
#     \begin{aligned}
#     \\
#     & \underset{\hspace{0.12cm} y_{ti},\hspace{0.12cm} z_{tkj}, \hspace{0.12cm} \psi_{tk}, \hspace{0.12cm} \gamma_{t}}{\text{maximize}} \quad \sum_{i \in \mathcal{I}} \frac{\sum_{t \in \mathcal{T}}    \Omega_i r_i x_{ti}}{d_i}\\
#     &\\
#     s.t. \quad & \sum_{t \in \mathcal{T}} y_{ti} \le 1 \hspace{11.5cm} \forall i \in \mathcal{I}\\
#     %r2
#     & \sum_{i \in \mathcal{I}} \frac{w_i x_{ti}}{d_i} \le W_{max} \hspace{10.2cm} \forall t \in \mathcal{T}\\
#     %r5
#     & x_{ti} \le d_i y_{ti} \hspace{11.7cm} \forall (t,i) \in \mathcal{T} \times \mathcal{I}\\
#     %r5
#     & \sum_{i \in \mathcal{I_S}} y_{ti}\theta_{si} \le (1-y_{ts})H \hspace{9.4cm} \forall (t,s) \in \mathcal{T} \times \mathcal{I_S}\\
#     %r6
#     & \sum_{k \in \mathcal{K}} l_k z_{tkj} \le L_{max}-\gamma_t \omega \hspace{9.3cm} \forall (t,j) \in \mathcal{T} \times \mathcal{J}\\
#     %rk
#     & \sum_{k \in \mathcal{K}} l_k \psi_{tk} \le \gamma_t B_{max} \hspace{10.1cm} \forall t \in \mathcal{T}\\
#     %r4
#     & \psi_{tk} + \sum_{j \in \mathcal{J}} z_{tkj}  \le 1-\epsilon+ \sum_{i \in \mathcal{I}(k)} \frac{x_{ti}(f_i +\phi_i m_i)}{d_i} \hspace{5.7cm} \forall (t,k) \in \mathcal{T} \times \mathcal{IBC}\\
#     %r3
#     & \sum_{i \in \mathcal{I}(k)} \frac{x_{ti}(f_i +\phi_i m_i)}{d_i} \le \psi_{tk} + \sum_{j \in \mathcal{J}} z_{tkj} \hspace{7.1cm} \forall (t,k) \in \mathcal{T} \times \mathcal{IBC}\\
#     %r4
#     & \psi_{tk} + \sum_{j \in \mathcal{J}} z_{tkj} \le 1-\epsilon + \sum_{i \in \mathcal{I}(k)} \frac{x_{ti}f_i}{d_i} + \sum_{\alpha \in \mathcal{H}(k)} \frac{x_{t\alpha} \phi_\alpha m_\alpha}{d_\alpha} \hspace{4.1cm} \forall (t,k) \in \mathcal{T} \times \mathcal{IND}\\
#     %r3
#     & \sum_{i \in \mathcal{I}(k)} \frac{x_{ti}f_i}{d_i} + \sum_{\alpha \in \mathcal{H}(k)} \frac{x_{t\alpha} \phi_\alpha m_\alpha}{d_\alpha} \le \psi_{tk} + \sum_{j \in \mathcal{J}} z_{tkj} \hspace{5.5cm} \forall (t,k) \in \mathcal{T} \times \mathcal{IND}\\
#     %might have to take out lambdas and infuse with IND
#     %r3
#     & \psi_{tk} + \sum_{j \in \mathcal{J}} z_{tkj} \le 1 - \epsilon + \sum_{i \in \mathcal{I}(k)} \frac{x_{ti}(f_i+\lambda m_i)}{d_i} + \sum_{\alpha \in \mathcal{L}(E^+)} \frac{x_{t\alpha} \lambda m_\alpha}{d_\alpha} + \sum_{\beta \in \mathcal{L}(\mathcal{IND}_{\neg f})} \frac{x_{t\beta} 2 m_\beta}{d_\beta} \hspace{0.5cm} \forall (t,k) \in \mathcal{T} \times \{\textit{E}\}\\
#     %r3
#     & \sum_{i \in \mathcal{I}(k)} \frac{x_{ti}(f_i+\lambda m_i)}{d_i} + \sum_{\alpha \in \mathcal{L}(E^+)} \frac{x_{t\alpha} \lambda m_\alpha}{d_\alpha} + \sum_{\beta \in \mathcal{L}(\mathcal{IND}_{\neg f})} \frac{x_{t\beta} 2 m_\beta}{d_\beta} \le \psi_{tk} + \sum_{j \in \mathcal{J}} z_{tkj} \hspace{1.8cm} \forall (t,k) \in \mathcal{T} \times \{\textit{E}\}\\
#     %r3
#     & \psi_{tk} + \sum_{j \in \mathcal{J}} z_{tkj} \le 1 - \epsilon + \sum_{i \in \mathcal{I}(k)} \frac{x_{ti}(f_i+\lambda m_i)}{d_i} + \sum_{\alpha \in \mathcal{L}(E_f^+)} \frac{x_{t\alpha} \lambda m_\alpha}{d_\alpha} + \sum_{\beta \in \mathcal{L}(\mathcal{IND}_{f})} \frac{x_{t\beta} 2 m_\beta}{d_\beta} \hspace{0.5cm} \forall (t,k) \in \mathcal{T} \times \{E_f\}\\
#     %r3
#     & \sum_{i \in \mathcal{I}(k)} \frac{x_{ti}(f_i+\lambda m_i)}{d_i} + \sum_{\alpha \in \mathcal{L}(E_f^+)} \frac{x_{t\alpha} \lambda m_\alpha}{d_\alpha} + \sum_{\beta \in \mathcal{L}(\mathcal{IND}_{f})} \frac{x_{t\beta} 2 m_\beta}{d_\beta} \le \psi_{tk} + \sum_{j \in \mathcal{J}} z_{tkj} \hspace{1.8cm} \forall (t,k) \in \mathcal{T} \times \{E_f\}\\
#     %r4
#     & \psi_{tk} + \sum_{j \in \mathcal{J}} z_{tkj} \le 1 - \epsilon + \sum_{i \in \mathcal{I}(k)} \frac{x_{ti}f_i}{d_i} \hspace{6.5cm} \forall (t,k) \in \mathcal{T} \times \{E^{+}, E_{f}^{+}\}\\
#     %r3
#     & \sum_{i \in \mathcal{I}(k)} \frac{x_{ti}f_i}{d_i} \le \psi_{tk} + \sum_{j \in \mathcal{J}} z_{tkj} \hspace{7.85cm} \forall (t,k) \in \mathcal{T} \times \{E^{+}, E_{f}^{+}\}\\
#     & z_{tkj} \in \mathbb{Z}_{\ge 0} \hspace{10.9cm} \forall (t,k,j) \in \mathcal{T} \times \mathcal{K} \times \mathcal{J}\\
#     & x_{ti} \in \mathbb{R}_{\ge 0} \hspace{11.8cm} \forall (t,i) \in \mathcal{T} \times \mathcal{I}\\
#     & y_{ti} \in \{0,1\} \hspace{11.5cm} \forall (t,i) \in \mathcal{T} \times \mathcal{I}\\
#     & \gamma_{t} \in \{0,1\} \hspace{11.6cm} \forall t \in \mathcal{T}\\
#     & \psi_{tk} \in \mathbb{Z}_{\ge 0} \hspace{11.75cm} \forall (t,k) \in \mathcal{T} \times \mathcal{K}\\
#     \end{aligned}
#     \end{equation*}
#     
#     
#     The core of what is going on down there. Consider extending to a 2-Stage Model for better last pallet efficiency. 
#     Details on documentation.

# In[108]:


class advanced_problem_parameters:
    
    weight_capacity=23500 #maximum weight a trailer can carry
    length_capacity=13200 #maximum length of stacked pallets length aka loading meters
    width_capacity=2400 #width of the truck
    
    def __init__(self, deliveries, dates, location, materials = materials, segregation_rules = smat, Lambda=1, max_days=7):
        #preprocess deliveries
        
        self.Lambda=Lambda
        #factor de penalización por mixed pallets e.g 1.15 adds 15% extra volume for mixed pallets as 
        #they're less efficient in terms of storage
        
        deliveries=deliveries.loc[deliveries["Name of sold-to party"]==location,:]
        #get STO lines for location of interest at current time
        deliveries=deliveries.reset_index().iloc[:,1:]
        self.Lead_Time=np.min(deliveries["Lead Times"])
        #get minimum lead time from observed lead times, which is the right one
        
        #plus a bunch of merging necessary because STOs have multiple due dates due to being shed back
        a=dates.drop_duplicates().reset_index().iloc[:,1:]
        b=dates.drop(columns="Del. Date").drop_duplicates().reset_index().iloc[:,1:]
        #delivery date will be assciated to the earliest date registered (highest reward) where material lines up 
        #with the reference document. We take the schedule line delivery date furthest in the 'past'
        
        for i in range(b.shape[0]):
            b.loc[i,"Delivery Date"]=np.max(a.loc[(a.loc[:,["Purch.Doc.","Material"
                ]]==b.loc[i,["Purch.Doc.","Material"]]).apply(lambda x: x[0]*x[1],axis=1),"Del. Date"])
            b.loc[i,"Late Delivery Date"]=np.min(a.loc[(a.loc[:,["Purch.Doc.","Material"
                ]]==b.loc[i,["Purch.Doc.","Material"]]).apply(lambda x: x[0]*x[1],axis=1),"Del. Date"])
        #Preparation for merging STO lines with their due delivery dates excluding the pushback dates. 
        #Such that we get one date, the right date, for each STO line
            
        new=pd.merge(b, deliveries ,how='inner',left_on=['Purch.Doc.','Material'
            ], right_on=['Purch.Doc.','Material']) #Merging due delivery dates to STO lines with stock availability
        
        
        new=pd.merge(new,materials[["Material","Pallet type Essers ","Breedte Essers ","Gross Weight"
                ,"Net Weight","GR slips Essers","Segregation group"]], how='inner', on="Material")
        #add relevant material specific data to STOs lines
        new=new.loc[new["GR slips Essers"]!=0,:]  #GR slips == 0 is an error that blows things up, placeholder
        new["Order Weight"]=new["Gross Weight"]*new['Delivery quantity']
        new["NPallets"]=new['Delivery quantity']/new["GR slips Essers"]
        #calculate total order weight + number of pallets per each order
        
        new["Days Until CDD"]=-new["Delivery Date"]
        new["Days Until CDD"]=new["Days Until CDD"]-self.Lead_Time
        #calculating number of days before comitted delivery date starting on current date
        new["Requested Delivery Date"]=(-new["Days Until CDD"]).apply(lambda x:(datetime.date.today()-datetime.timedelta(days=x)).strftime('%d.%m.%Y'))
        #formatting CDD in terms of strf
        
        new=new.loc[new.loc[:,"Days Until CDD"]<=(max_days),:].reset_index().iloc[:,1:]
        #filtering out STO lines whose CDD is greater or equal than max_days
        
        
        self.outside_consideration=new.loc[new.loc[:,"Days Until CDD"]>(max_days+self.Lead_Time),:]
        #storing STO lines that are present but not considered on current timeframe (too far ahead into the future)
        
        
        self.pallet=materials[["Pallet type Essers ","Breedte Essers "]].drop_duplicates().dropna().reset_index().iloc[:,1:]
        self.pallet['Pallet type Essers ']=[i+str('f') for i in self.pallet['Pallet type Essers ']]
        self.pallet=pd.concat([materials[["Pallet type Essers ","Breedte Essers "]].drop_duplicates().dropna().reset_index().iloc[:,1:],self.pallet], ignore_index=True)
        #creating data frame containing pallet types and their dimension when stored orthogonal respect to the truck (shorter side)
        #pallets containing flammables are considered special pallet types as flammable products for practical purposes (see documentation)
        
        x=new.loc[np.isin(new.loc[:,'Segregation group'],['2.1','4.1 UN3175']),:]['Pallet type Essers '] #'3Ac','3Al','3N','4.1 UN3238','4.2'
        new.loc[x.index,'Pallet type Essers ']=[i+str('f') for i in x]
        #changing pallet label to flammable SKUs such that we segregate flammables into different pallets
        
        self.pallet=self.pallet.loc[self.pallet.loc[:,"Pallet type Essers "].isin(new["Pallet type Essers "].unique())].reset_index().iloc[:,1:]
        #dropping redundant pallet types
        
        #function that splits large STO lines, such that they can be distributed among various trucks
        #optimal split is selected between LMs and weight. 
        def advanced_splitter(og_new): 
    
                #split item orders that exceed volume in terms of pallets
                to_split_by_weight=list()
                to_split_by_length=list()
                weight_capacity=23500
                length_capacity=2*13200
                df=og_new
                
                for i in range(df.shape[0]):
                    #get LMs, weight ratios for all orders. Split the selected order as defined by the bigger ratio assuming
                    #it is greater than 1. Next up. 
                    
                    weight_capacity_ratio=df.loc[i,"Order Weight"]/weight_capacity
                    length_capacity_ratio=df.loc[i,"NPallets"]/np.ceil(length_capacity/int(self.pallet.loc[self.pallet.loc[:,
                    'Pallet type Essers ']==df.loc[i,"Pallet type Essers "]]["Breedte Essers "]))
                    
                    if weight_capacity_ratio>1 or length_capacity_ratio>1:
                    #identify orders of too many pallets
                        if length_capacity_ratio>=weight_capacity_ratio:
                            
                            to_split_by_length.append(i)
                            
                        else:
                            
                            to_split_by_weight.append(i)
                            #take index of orders to be split according to split type
                            
                
                def split_by_length(new, to_split):

                    df=new
                    df=df.loc[to_split,:]
                    biggies=pd.DataFrame()
                    for i in to_split:
                        load=df.loc[i,"NPallets"]
                        cap=np.floor(length_capacity/self.pallet.loc[self.pallet.loc[:,'Pallet type Essers ']
                        ==df.loc[i,"Pallet type Essers "]]["Breedte Essers "]) #number of pallets per truck at maximum
                        while load>cap:
                            load-=cap
                            row=pd.DataFrame(df.loc[i,:]).transpose()
                            row['Delivery quantity']=row["GR slips Essers"]*cap # does the split of an item order to whichever are needed
                            biggies=pd.concat([biggies,row], ignore_index=True) # stacks item splits on top of each other
                            #successively split STO line on lines with NPallets equal to maximum number of pallets per truck
                            if load<=cap:
                            #residual pallets with which we can't fill an entire truck go into last STO line
                                row['Delivery quantity']=load*row["GR slips Essers"]
                                biggies=pd.concat([biggies,row], ignore_index=True)

                    new=new.drop(to_split)
                    #drop STO lines that have been split
                    new=pd.concat([new,biggies], ignore_index=True).reset_index().iloc[:,1:]
                    new["Order Weight"]=new["Gross Weight"]*new['Delivery quantity']
                    new["NPallets"]=new['Delivery quantity']/new["GR slips Essers"]

                    return new
                
                
                #splits item orders that are too heavy for one truck
                #exact same resoning as below applied to max number of pallets a truck can carry in terms of weight
                def split_by_weight(new, to_split):
                    
                    df=new
                    df=df.loc[to_split,:]
                    biggies=pd.DataFrame()
                    for i in to_split:
                        load=df.loc[i,"Order Weight"]
                        weight_per_pallet=df.loc[i,"Gross Weight"]*df.loc[i,"GR slips Essers"]
                        cap=weight_capacity
                        #cap=weight_capacity #-1 substracting one or 2 can make sure that pallets fit
                        while load>cap:

                            x=np.floor(weight_capacity/weight_per_pallet)*df.loc[i,"GR slips Essers"]
                            #instead of using pallets as reference, we use units for historical reasons. It's the same thing.

                            load-=x*df.loc[i,"Gross Weight"]
                            row=pd.DataFrame(df.loc[i,:]).transpose()
                            row['Delivery quantity']=x
                            row['Order Weight']=x*df.loc[i,"Gross Weight"]
                            biggies=pd.concat([biggies,row], ignore_index=True)
                            #split STO lines into multiple lines such that they carry the maximum number of full pallets
                            #that could be feasibly carried by a truck, with weight as the limiting factor
                            if load<=cap:
                                row['Delivery quantity']=np.round(load/df.loc[i,"Gross Weight"])
                                row['Order Weight']=row['Delivery quantity']*df.loc[i,"Gross Weight"]
                                biggies=pd.concat([biggies,row], ignore_index=True)

                    new=new.drop(to_split)
                    new=pd.concat([new,biggies], ignore_index=True).reset_index().iloc[:,1:]
                    #readjust all columns according to changes
                    new["Order Weight"]=new["Gross Weight"]*new['Delivery quantity']
                    new["NPallets"]=new['Delivery quantity']/new["GR slips Essers"]

                    return new
            
                og_new=split_by_weight(og_new, to_split_by_weight)
                og_new=split_by_length(og_new, to_split_by_length)

                return og_new
        
        new=advanced_splitter(new)
        new=advanced_splitter(new)
        
        new.loc[new.loc[:,"DPrio"]==55,"Priority Level"]=100
        new.loc[new.loc[:,"DPrio"]==65,"Priority Level"]=1
        new.loc[new.loc[:,"DPrio"]==60,"Priority Level"]=0.01
        #assigning multipliers to STO lines according to priority
    
        new["Late Delivery Date"]=new['Late Delivery Date'].apply(lambda x:(datetime.date.today()-datetime.timedelta(days=x)).strftime('%d.%m.%Y'))
        #late delivery date just for FYI
        
        new["Delivery Date"]=100/(1+np.exp(1)**(0.05*(-5*new["Delivery Date"])))*new['NPallets']*new["Priority Level"]
        #applying function to assign value in terms of CDD, priority and controlling for volume (see documentation)
    
        #next up we start preparing the relevant data for being inputted into optimization model
        nP=new.shape[0]
        products=set(range(0,nP))
        #define set of STO lines
        
        
        pallet_types=set(new["Pallet type Essers "].unique())
        #define set of relevant pallet types
        
        self.length = {}
        for i in range(0,len(pallet_types)):
            self.length[self.pallet.loc[i,"Pallet type Essers "]] = self.pallet.loc[i,"Breedte Essers "]
        #define parameter of pallet dimesion over pallet types
            
        print(self.length)
        self.Vm=defaultdict(set)
        for i in pallet_types:
            self.Vm[i]=set(new[new["Pallet type Essers "]==i].index) 
        #set of sets of STO lines' indices according to the pallet type they belong to
            
        
        #defining a bunch of parameters next up
        self.reward = {}
        self.weight = {}
        self.demand = {}
        self.units_per_pallet = {}
        self.full_pallets = {}
        self.mixed_pallets = {}
        for i in products:
            self.reward[i] = new.loc[i,"Delivery Date"] #reward per shipping ith STO line
            self.weight[i] = new.loc[i,"Order Weight"] #weight of ith STO line
            self.demand[i] = new.loc[i,'Delivery quantity'] #units requested of ith STO line's SKU
            self.units_per_pallet[i] = new.loc[i,'GR slips Essers'] #units per pallet for SKU at ith STO line
            self.full_pallets[i]=np.floor(new.loc[i,'NPallets']) #full pallets at ith STO line
            self.mixed_pallets[i]=new.loc[i,'NPallets']-np.floor(new.loc[i,'NPallets']) #Fractional pallets as ith STO line
            if ((new.loc[i,"Pallet type Essers "]=="IND") | (new.loc[i,"Pallet type Essers "]=="IND+")) & (self.mixed_pallets[i]>0.75):
                    self.mixed_pallets[i]=0
                    self.full_pallets[i]+=1
            if ((new.loc[i,"Pallet type Essers "]=="INDf") | (new.loc[i,"Pallet type Essers "]=="IND+f")) & (self.mixed_pallets[i]>0.75):
                    self.mixed_pallets[i]=0
                    self.full_pallets[i]+=1
            #industrial pallet types over 75% capacity are considered full for practical purposes

        self.down_Set=defaultdict(set)
        for i in ["EUR+","IND","IND+","EUR+f","INDf","IND+f"]:
            if i=="EUR+" or i=="EUR+f":
                self.down_Set[i]=set(new[(new["Pallet type Essers "]==i) & (np.fromiter(self.mixed_pallets.values(), dtype=float)<1) & (np.fromiter(self.mixed_pallets.values(), dtype=float)>0)].index)#((new['NPallets']-np.floor(new['NPallets']))<=0.3)].index)
            else:
                self.down_Set[i]=set(new[(new["Pallet type Essers "]==i) & (np.fromiter(self.mixed_pallets.values(), dtype=float)<=0.75) & (np.fromiter(self.mixed_pallets.values(), dtype=float)>0)].index)#((new['NPallets']-np.floor(new['NPallets']))<=0.5)].index)
        print(self.down_Set)
        #set of sets where indexes of pallet types where mixed pallets will be stored on Europallets despite it not being its default pallet type
        
        self.up_Set=defaultdict(set)
        #set of sets that stores indexes of orders with almost full mixed pallets that still can take up more volume
        #not applicable at BE04 so should be always null. Thus there's an unrealistic parameter (999) to initialize the set 
        for i in ["EUR+","EUR+f"]:
                self.up_Set[i]=set(new[(new["Pallet type Essers "]==i) & (np.fromiter(self.mixed_pallets.values(), dtype=float)==999)].index)
        #print(self.down_Set)
        #print(self.up_Set)
            
        
        #get potentially hazardous products
        segregate=new.loc[(new["Segregation group"]!="None") & (new["Segregation group"]!="NH"),:]
        #get STO lines to segregate
        segregated_products=segregate.index
        segregate=segregate.reset_index().iloc[:,1:]
        
        #make segregation matrix from item order x item order. Called Theta in documentation. 
        #The imported segregation matrix was classifiction type x classification type
        segregation_matrix = pd.DataFrame(np.zeros((len(segregate.index),len(segregate.index))))
        for i in range(0,len(segregate.index)):
            for j in range(0,len(segregate.index)):
                segregation_matrix.loc[i,j]=segregation_rules.loc[segregate.loc[i,"Segregation group"],segregate.loc[j,"Segregation group"]]
                #creating new conflict matrix for cartesian product of STO lines, instead of that of segregation groups which we get as smat
        segregation_matrix.index=segregated_products
        segregation_matrix.columns=segregated_products
        segregation_matrix=segregation_matrix.loc[segregation_matrix.apply(sum,axis=1)>0,segregation_matrix.apply(sum,axis=0)>0]
        #removing STO lines that present no segregation conflicts in practice
        self.location=location
        
        self.segregate_in_practice=False
        if segregation_matrix.shape[0]>0:
        #if empirical segregation matrix is not null then activate segregation
            self.segregation_encoding=segregation_matrix.stack().to_dict()
            self.segregated_products=set(segregation_matrix.index)
            self.segregation_matrix=segregation_matrix
            self.segregate_in_practice=True
        #will allow us to deactivate segregation were to be applied on paper but were not necessary in practice
    
        self.pallet_types=pallet_types
        self.products=products
    
        self.data=new
        
        self.table=new.pivot_table(index='Segregation group', columns='Pallet type Essers ', values='NPallets', aggfunc='sum')
    
    def optimize(self, max_time=60, tol_weight=0.95, tol_length=0.97, MILP_solver='cbc', to_segregate=False, to_display=False, max_trucks=20):
        
        startTotalTime = time.time()
        keep_going=True
        nT=1

        H=100000 # should be greater than number of STOs, just a big number
        omega=1200 #milimeters spare in the truck in order to rotate pallets
        solver = po.SolverFactory(MILP_solver)
        model="None"
        result="None"
        #"none"s are necessary preparation as we will usually wish to store previous to last solution found by virtue of the fact 
        #that we are systematically running into a wall to then fall back into our sought after solution (more on documentation)
        while (keep_going==True):

            trucks=set(range(0,nT))
            if nT>1:
                #storing solutions found at previous iteration
                chosen_model=model
                chosen_result=result
                
            model=pe.ConcreteModel()

            #initializing various sets over which parameters, variables and restrictions lie.
            model.trucks=pe.Set(initialize=trucks)
            model.products=pe.Set(initialize=self.products)
            model.pallet_types=pe.Set(initialize=self.pallet_types)
            model.lanes=pe.Set(initialize={0,1})

            #specify parameters over relevant sets
            model.reward=pe.Param(model.products,initialize=self.reward)
            model.weight=pe.Param(model.products,initialize=self.weight)
            model.length=pe.Param(model.pallet_types,initialize=self.length)
            model.full_pallets=pe.Param(model.products,initialize=self.full_pallets)
            model.mixed_pallets=pe.Param(model.products,initialize=self.mixed_pallets)
            
            model.demand=pe.Param(model.products,initialize=self.demand)
            model.units_per_pallet=pe.Param(model.products,initialize=self.units_per_pallet)
            if to_segregate==True:
                model.segregated_products=pe.Set(initialize=self.segregated_products)
                model.segregation_encoding=pe.Param(self.segregated_products,self.segregated_products,
                                                    initialize=self.segregation_encoding)
                
            model.Vm=pe.Param(model.pallet_types, initialize=self.Vm, default=set(), within=pe.Any)
            
            #initialize set of sets of non-standard europallets to be stored in europallets
            initialize_down_Set=False
            down_Set_nonempty=self.down_Set
            for i in ["EUR+","IND+","IND","EUR+f","IND+f","INDf"]:
                if bool(self.down_Set[i])==True:
                    initialize_down_Set=True
                else:
                    down_Set_nonempty.pop(i)
                    
            if initialize_down_Set==True:
                model.down_Set=pe.Param(model.pallet_types, initialize=down_Set_nonempty, default=set(), within=pe.Any)
            
            initialize_up_Set=False
            up_Set_nonempty=self.up_Set
            for i in ["EUR+","EUR+f"]:
                if bool(self.up_Set[i])==True:
                    initialize_up_Set=True
                else:
                    up_Set_nonempty.pop(i)
                    
            if initialize_up_Set==True:
                model.up_Set=pe.Param(model.pallet_types, initialize=up_Set_nonempty, default=set(), within=pe.Any)
            
            time.sleep(2)
            #initializing variables over sets
            model.z=pe.Var(model.trucks,model.pallet_types,model.lanes,domain=pe.NonNegativeIntegers,initialize=0)
            model.psi=pe.Var(model.trucks,model.pallet_types,domain=pe.NonNegativeIntegers,initialize=0)
            model.alpha=pe.Var(model.trucks,domain=pe.Binary,initialize=0)
            model.y=pe.Var(model.trucks,model.products,domain=pe.Binary,initialize=0)
            model.x=pe.Var(model.trucks,model.products,domain=pe.NonNegativeReals,initialize=0)
            

            expr=sum(model.reward[i]*model.x[t,i]/model.demand[i] for i in model.products for t in model.trucks)
            #objective function
            model.obj=pe.Objective(sense=pe.maximize, expr=expr)
                    
            #constraints
            #ensure trucks do not exceed weight capacity
            #dummy constraints are redundant and have the purpose of comfortable data retrieval

            #ensure an sto line is served only once
            def unicity(model,i):
                constraint=(sum(model.y[t,i] for t in model.trucks)<=1)
                return constraint

            def dummy_length_capacity(model,t,j):
                covered=sum(model.z[t,k,j]*model.length[k] for k in model.pallet_types)
                constraint=(covered<=advanced_problem_parameters.length_capacity)
                return constraint

            #ensure loading meters are under a certain threshold
            def length_capacity(model,t,j):
                covered=sum(model.z[t,k,j]*model.length[k] for k in model.pallet_types)
                constraint=(covered<=advanced_problem_parameters.length_capacity-model.alpha[t]*omega)
                return constraint

            #ensure weight capacity is under certain threshold
            def weight_capacity(model,t):
                covered=sum(model.x[t,i]/model.demand[i]*model.weight[i] for i in model.products)
                constraint=(advanced_problem_parameters.weight_capacity>=covered)
                return constraint

            #ensure we do not ship more than demanded and choice of SKUs complies with segregation
            def semi_integrality(model,t,i):
                constraint=(model.x[t,i]<=model.demand[i]*model.y[t,i])
                return constraint 

            #determine pallets such that they house equal or bigger than the volume they carry
            #for more detail go to the more formal documentation
            def integer_pallets_low(model,t,k):

                if k=="EUR+" or k=="EUR+f":
                    covered_full=sum(model.x[t,ik]/model.demand[ik]*model.full_pallets[ik] for ik in model.Vm[k])
                    if initialize_up_Set==True:
                        if bool(self.up_Set[k])==True:
                            covered_residue=sum(model.x[t,ik]/model.demand[ik]*self.Lambda*model.mixed_pallets[ik] for ik in model.up_Set[k])
                        else:
                            covered_residue=0
                    else:
                        covered_residue=0
                    constraint=(covered_full+covered_residue)<=(model.psi[t,k]+sum(model.z[t,k,j] for j in model.lanes))
                    return constraint

                #change to complementary residue set
                elif k=="IND" or k=="IND+" or k=="INDf" or k=="IND+f":
                    covered_full=sum(model.x[t,ik]/model.demand[ik]*model.full_pallets[ik] for ik in model.Vm[k])
                    constraint=(covered_full<=(model.psi[t,k]+sum(model.z[t,k,j] for j in model.lanes)))
                    return constraint

                elif k=="EUR":
                    covered_eur=sum(model.x[t,ik]/model.demand[ik]*(model.full_pallets[ik]+self.Lambda*model.mixed_pallets[ik]) for ik in model.Vm[k])

                    if initialize_down_Set==True:
                        #there's a better way to do this, make set of non empty sets
                        if bool(self.down_Set["EUR+"])==True:
                            covered_eurplus=sum(model.x[t,ik]/model.demand[ik]*self.Lambda*model.mixed_pallets[ik] for ik in model.down_Set["EUR+"])
                        else:
                            covered_eurplus=0
                        if bool(self.down_Set["IND"])==True:
                            covered_ind=sum(model.x[t,ik]/model.demand[ik]*2*model.mixed_pallets[ik] for ik in model.down_Set["IND"])
                        else:
                            covered_ind=0
                        if bool(self.down_Set["IND+"])==True:
                            covered_indplus=sum(model.x[t,ik]/model.demand[ik]*2*model.mixed_pallets[ik] for ik in model.down_Set["IND+"])
                        else:
                            covered_indplus=0

                    else:

                        covered_eurplus=0
                        covered_ind=0
                        covered_indplus=0

                    constraint=(covered_eur+covered_eurplus+covered_ind+covered_indplus<=(model.psi[t,k]+sum(model.z[t,k,j] for j in model.lanes)))
                    return constraint

                elif k=="EURf":
                    covered_eur=sum(model.x[t,ik]/model.demand[ik]*(model.full_pallets[ik]+self.Lambda*model.mixed_pallets[ik]) for ik in model.Vm[k])

                    if initialize_down_Set==True:
                        #there's a better way to do this, make set of non empty sets
                        if bool(self.down_Set["EUR+f"])==True:
                            covered_eurplus=sum(model.x[t,ik]/model.demand[ik]*self.Lambda*model.mixed_pallets[ik] for ik in model.down_Set["EUR+f"])
                        else:
                            covered_eurplus=0
                        if bool(self.down_Set["INDf"])==True:
                            covered_ind=sum(model.x[t,ik]/model.demand[ik]*2*model.mixed_pallets[ik] for ik in model.down_Set["INDf"])
                        else:
                            covered_ind=0
                        if bool(self.down_Set["IND+f"])==True:
                            covered_indplus=sum(model.x[t,ik]/model.demand[ik]*2*model.mixed_pallets[ik] for ik in model.down_Set["IND+f"])
                        else:
                            covered_indplus=0

                    else:

                        covered_eurplus=0
                        covered_ind=0
                        covered_indplus=0

                    constraint=(covered_eur+covered_eurplus+covered_ind+covered_indplus<=(model.psi[t,k]+sum(model.z[t,k,j] for j in model.lanes)))
                    return constraint

                elif k=="COG" or k=="COK" or k=="COGf" or k=="COKf":

                    covered=sum(model.x[t,ik]/model.demand[ik]*(model.full_pallets[ik]+self.Lambda*model.mixed_pallets[ik]) for ik in model.Vm[k])
                    constraint=(covered<=(model.psi[t,k]+sum(model.z[t,k,j] for j in model.lanes)))
                    return constraint

            #determine pallets such that we do not pich more than we need. Volume SKUs+0.99 \leq Volume in Pallets
            def integer_pallets_high(model,t,k):

                if k=="EUR+" or k=="EUR+f":
                    covered_full=sum(model.x[t,ik]/model.demand[ik]*model.full_pallets[ik] for ik in model.Vm[k])
                    if initialize_up_Set==True:
                        if bool(self.up_Set[k])==True:
                            covered_residue=sum(model.x[t,ik]/model.demand[ik]*self.Lambda*model.mixed_pallets[ik] for ik in model.up_Set[k])
                        else:
                            covered_residue=0
                    else:
                        covered_residue=0
                    constraint=(0.99+covered_full+covered_residue)>=(model.psi[t,k]+sum(model.z[t,k,j] for j in model.lanes))
                    return constraint
                #change to complementary residue set

                elif k=="IND" or k=="IND+" or k=="INDf" or k=="IND+f":
                    covered_full=sum(model.x[t,ik]/model.demand[ik]*model.full_pallets[ik] for ik in model.Vm[k])
                    constraint=(0.99+covered_full>=(model.psi[t,k]+sum(model.z[t,k,j] for j in model.lanes)))
                    return constraint

                elif k=="EUR":
                    covered_eur=sum(model.x[t,ik]/model.demand[ik]*(model.full_pallets[ik]+self.Lambda*model.mixed_pallets[ik]) for ik in model.Vm[k])

                    if initialize_down_Set==True:
                        #there's a better way to do this, make set of non empty sets
                        if bool(self.down_Set["EUR+"])==True:
                            covered_eurplus=sum(model.x[t,ik]/model.demand[ik]*self.Lambda*model.mixed_pallets[ik] for ik in model.down_Set["EUR+"])
                        else:
                            covered_eurplus=0
                        if bool(self.down_Set["IND"])==True:
                            covered_ind=sum(model.x[t,ik]/model.demand[ik]*2*model.mixed_pallets[ik] for ik in model.down_Set["IND"])
                        else:
                            covered_ind=0
                        if bool(self.down_Set["IND+"])==True:
                            covered_indplus=sum(model.x[t,ik]/model.demand[ik]*2*model.mixed_pallets[ik] for ik in model.down_Set["IND+"])
                        else:
                            covered_indplus=0

                    else:

                        covered_eurplus=0
                        covered_ind=0
                        covered_indplus=0

                    constraint=(0.99+covered_eur+covered_eurplus+covered_ind+covered_indplus>=(model.psi[t,k]+sum(model.z[t,k,j] for j in model.lanes)))
                    return constraint

                elif k=="EURf":
                    covered_eurf=sum(model.x[t,ik]/model.demand[ik]*(model.full_pallets[ik]+self.Lambda*model.mixed_pallets[ik]) for ik in model.Vm[k])

                    if initialize_down_Set==True:
                        #there's a better way to do this, make set of non empty sets
                        if bool(self.down_Set["EUR+f"])==True:
                            covered_eurplusf=sum(model.x[t,ik]/model.demand[ik]*self.Lambda*model.mixed_pallets[ik] for ik in model.down_Set["EUR+f"])
                        else:
                            covered_eurplusf=0
                        if bool(self.down_Set["INDf"])==True:
                            covered_indf=sum(model.x[t,ik]/model.demand[ik]*2*model.mixed_pallets[ik] for ik in model.down_Set["INDf"])
                        else:
                            covered_indf=0
                        if bool(self.down_Set["IND+f"])==True:
                            covered_indplusf=sum(model.x[t,ik]/model.demand[ik]*2*model.mixed_pallets[ik] for ik in model.down_Set["IND+f"])
                        else:
                            covered_indplusf=0

                    else:

                        covered_eurplusf=0
                        covered_indf=0
                        covered_indplusf=0

                    constraint=(0.99+covered_eurf+covered_eurplusf+covered_indf+covered_indplusf>=(model.psi[t,k]+sum(model.z[t,k,j] for j in model.lanes)))
                    return constraint

                elif k=="COG" or k=="COK" or k=="COGf" or k=="COKf":

                    covered=sum(model.x[t,ik]/model.demand[ik]*(model.full_pallets[ik]+self.Lambda*model.mixed_pallets[ik]) for ik in model.Vm[k])
                    constraint=(0.99+covered)>=(model.psi[t,k]+sum(model.z[t,k,j] for j in model.lanes))
                    return constraint

            #makes sure pallets fit in terms of width
            def longitudinal_length_capacity(model,t):
                covered=sum(model.psi[t,k]*model.length[k] for k in model.pallet_types)
                constraint=(covered<=model.alpha[t]*advanced_problem_parameters.width_capacity)
                return constraint

            def dummy_longitudinal_length_capacity(model,t):
                covered=sum(model.psi[t,k]*model.length[k] for k in model.pallet_types)
                constraint=(covered<=advanced_problem_parameters.width_capacity)
                return constraint
                
            #initializing constraints up next. 
            model.con_unicity=pe.Constraint(model.products, rule=unicity)
            model.con_semi_integrality=pe.Constraint(model.trucks,model.products, rule=semi_integrality)
            model.con_length_capacity=pe.Constraint(model.trucks,model.lanes, rule=length_capacity)
            model.con_dummy_length_capacity=pe.Constraint(model.trucks,model.lanes, rule=dummy_length_capacity)
            model.con_longitudinal_length_capacity=pe.Constraint(model.trucks, rule=longitudinal_length_capacity)
            model.con_dummy_longitudinal_length_capacity=pe.Constraint(model.trucks, rule=dummy_longitudinal_length_capacity)
            model.con_weight_capacity=pe.Constraint(model.trucks, rule=weight_capacity)
            model.con_integer_pallets_low=pe.Constraint(model.trucks, model.pallet_types, rule=integer_pallets_low)
            model.con_integer_pallets_high=pe.Constraint(model.trucks, model.pallet_types, rule=integer_pallets_high)
    
            
            #only add segregation if pre-specified
            if to_segregate==True:
                
                #ensure there is no segregation conflicts
                def segregation(model,t,s):
                    covered=sum(sum(model.y[t,i]*model.segregation_encoding[s,i] for i in model.segregated_products) for j in model.lanes)
                    constraint=covered<=(1-model.y[t,s])*H
                    return constraint
                                                         
                model.con_segregation=pe.Constraint(model.trucks, model.segregated_products, rule=segregation)

            startTime = time.time()

            result = solver.solve(model, timelimit=max_time)#, tee=True, keepfiles=True, logfile=str(location)+str(nT)+".log")
            #print(result)
            time.sleep(3)

            executionTime = (time.time() - startTime)
            
                
            #print(result)
            if to_display==True:
                print(model.display())


            print('Iteration ' + str(nT) + ' | Execution time in seconds: ' + str(np.round(executionTime,3)) + 
                  ' | Objective value: ' + str(result['Problem'][0]["Lower bound"]) + ' | Worst-case suboptimality: ' + 
                  str(np.round(100*(result['Problem'][0]["Upper bound"]/result['Problem'][0]["Lower bound"]-1),5)) + '%')

            #print(model.alpha.extract_values())
            con_weight=np.zeros(nT)
            con_length=np.zeros(nT)
            for i in trucks:
                con_weight[i]=model.con_weight_capacity[i].body()
                #getting weight per truck
                con_length[i]=model.con_dummy_longitudinal_length_capacity[i].body()
                for l in range(0,2):
                    con_length[i]+=model.con_dummy_length_capacity[i,l].body()
                    #getting LMs per truck

            print(con_length/2)
            print(con_weight)

            '''print(model.alpha.display())
            print(model.psi.display())
            print(model.con_longitudinal_length_capacity.display())'''
        
            if (np.sum((con_weight>(advanced_problem_parameters.weight_capacity*tol_weight)) | (con_length>(2*advanced_problem_parameters.length_capacity*tol_length)))<nT) | (max_trucks==nT):
            #if all trucks are either not full at x% capacity via length or weight, or we got to the maximum possible number of 
            #trucks we come to a final solution and generate output. Else we try adding one truck more.
            
                
                keep_going=False
                #major while loop stops


                if (nT==1) & (max_trucks>1):
                #if we fail to fill one truck we provide as output a class notifying of this fact
                #class mimics intended output class in terms instane variables to avoid uninformative errors

                    class solution:
                        def __init__(self):
                            self.choice_of_trucks=0
                            self.overview='Not ready for shipment.'
                            self.solution='Not ready for shipment.'
                            self.output='Not ready for shipment.'
                            self.leftout='Not ready for shipment.'
                            self.leftout_table='Not ready for shipment.'
                            self.execution_time='Not ready for shipment.'
                            #self.conflicts='Not ready for shipment.'
                            
                    print("STOP: Not ready for shipment.")
                    print('')
                    return solution()

                else:

                    if (max_trucks==nT) & (np.sum((con_weight>(advanced_problem_parameters.weight_capacity*tol_weight)) | (con_length>(2*advanced_problem_parameters.length_capacity*tol_length)))==nT):
                    #if we got stopped because we hit the maximum number of trucks, we stop at the last iteration
                    #else we will settle at solution previous to last
                        chosen_model=model
                        chosen_result=result
                        nT=nT+1
            
                    print("STOP: Optimum at iteration " + str(nT-1) + ' | Total time: ' + str((time.time() - startTotalTime)))
                    print('')

                    output_y=pd.DataFrame.from_dict(chosen_model.y.extract_values(), orient='index', columns=[str(chosen_model.y)])
                    output_y=output_y.reset_index()
                    #getting solution to binary variables representing order activation

                    output_y['STO']=0
                    output_y['Truck']=0
                    for i in range(0,output_y.shape[0]):
                        output_y.loc[i,'Truck']=output_y['index'][i][0]+1
                        output_y.loc[i,'STO']=output_y['index'][i][1]
                    #building simple data frame to store found solution on a friendly format

                    output_x = pd.DataFrame.from_dict(chosen_model.x.extract_values(), orient='index', columns=[str(chosen_model.x)])
                    output_x=output_x.reset_index()
                    output_x['STO']=0
                    output_x['Truck']=0
                    for i in range(0,output_x.shape[0]):
                        output_x.loc[i,'Truck']=output_x['index'][i][0]+1
                        output_x.loc[i,'STO']=output_x['index'][i][1]
                    #same as previous with real non-negative variables whose movement range is defined by activation variables


                    output_x=output_x.drop('index',axis=1)                     
                    output_y=output_y.drop('index',axis=1)

                    output_y=pd.merge(output_y,output_x, how='inner', on=["Truck","STO"])

                    leftout=self.data.loc[(output_y.pivot_table(index='STO', values='x', aggfunc='sum')['x']==0),:]
                    #storing orders we did not serve
                    output_y=output_y.loc[output_y["x"]>0,:].drop('y',axis=1)
                    #storing served orders, even if partially served.

                    check=pd.merge(output_y,self.data,how="left", left_on='STO', right_index=True).reset_index().drop('index',axis=1)
                    output=pd.merge(output_y,self.data.loc[:,["Name of sold-to party","Purch.Doc.",'Description','Material', 
                                     'Delivery quantity','Delivery Date','NPallets', "Days Until CDD",'Pallet type Essers ', 'Segregation group','Order Weight','Requested Delivery Date','Late Delivery Date',"DPrio","DlvTy",'Deliv.date',
                    'Plan GI Dt', 'Lead Times']],
                                    how="left", left_on='STO', right_index=True).reset_index().drop('index',axis=1).drop(columns="STO")
                    output=output.rename(columns={"x":"Quantity Delivered"})
                    output=output.rename(columns={"Delivery Date":"Reward"})
                    #keep condensing output in a friendly format that includes each order's relevant material data

                    output=output.loc[:,["Truck","Name of sold-to party","Purch.Doc.",'Description',"DPrio",'Material','Quantity Delivered',
                                   'Delivery quantity','NPallets', 'Reward', "Days Until CDD",'Pallet type Essers ', 'Segregation group','Order Weight','Requested Delivery Date','Late Delivery Date',"DlvTy",'Deliv.date',
                    'Plan GI Dt', 'Lead Times']]
                    #sorting column order


                    '''if to_segregate==True:
                    #serves to verify that segregation is done properly, if necessary (superflous)
                        check=check.pivot_table(index='Segregation group', columns='Truck', values='NPallets', aggfunc='sum')
                        conflicts=np.zeros(len(check.columns))
                        for j in check.columns:
                            for i in check.index[check.index!='None']:
                                for k in check.index[check.index!='None']:
                                    if(check.loc[i,j]>0) & (check.loc[k,j]>0):
                                        if smat.loc[i,k]==1:
                                            conflicts[j]=conflicts[j]+1

                        conflicts=np.sum(conflicts)

                    else:

                        conflicts=0'''
                        

                    class solution:
                    #storing solution in a class

                        def __init__(self,chosen_result,chosen_model,output,nT):
                            self.choice_of_trucks=nT-1 #number of trucks 
                            self.overview=chosen_result #statistics relating to solution and sub-optimality were it present
                            self.solution=chosen_model #raw solution. Is to be accessed with self.solution.display()
                            self.output=output #friendly version of the output
                            self.leftout=leftout #storing left out orders
                            self.leftout_table=pd.pivot_table(leftout, values='NPallets', index='Name of sold-to party', columns='Pallet type Essers ', aggfunc='sum')
                            #quick overview of unassigned volumes
                            #self.conflicts=conflicts #gives number of conflicts between orders per truck. Should be 0 without exception
                            self.execution_time=(time.time() - startTotalTime)

                    return solution(chosen_result, chosen_model, output, nT)
            else:

                nT+=1


# In[109]:


def execute_optimizer(deliveries, dates, materials=materials, max_time=120, 
                      tol_weight=0.95, tol_length=0.97, MILP_solver='cbc', Lambda=1.1, max_trucks=3, max_days=21):
#wrapper function to execute optimization for all considered locations while specifying where segregation is to be active and 
#condensing all location's outputs into a single one
    
    startTime = time.time()
    items_df=pd.DataFrame() #to store STO lines to be carried per truck
    trucks_df=pd.DataFrame() #to store truck overview
    residual_deliveries=pd.DataFrame() #to store unassigned STO lines
    weights=list() #store weight per truck
    lengths=list() #store loading meters per truck


    #iterates for all locations and saves outputs
    for i in np.unique(deliveries["Name of sold-to party"]):

        if np.isin(i,['FI01 Ecolab Europe GmbH','GB03_Trafford Park','NO01Norway DC','IE01_Mullingar Factory/DC','PT01 SANTIAGO DO VELHOS'])==True:
        #activate segregation if locations is in list
            Active_Segregation=True

        else:

            Active_Segregation=False

        print('Location: ' + i)
        startTime2 = time.time()
        z=advanced_problem_parameters(deliveries, dates, i, materials, Lambda=Lambda, max_days=max_days)
        #get all parameters to be considered for each optimization problem, of which there is one per location
        residual_deliveries=pd.concat([residual_deliveries,
                z.outside_consideration.loc[:,["Material","Description",
                "Name of sold-to party","Purch.Doc.","Delivery quantity","DPrio","DlvTy","Days Until CDD", 'Deliv.date',
                'Plan GI Dt', 'Lead Times']]], ignore_index=True)
        #get all STO lines outside timeframe considered
    
        print('Iteration 0 | Initialization time:', time.time()-startTime2)
        
        if z.data.shape[0]==0:
        #if there's only STO lines outside observed timeframe move on to next destination              
                print('STOP: Not volume on considered time frame.')
                print('')
                continue

        try:
            
            a=z.optimize(max_time=max_time, tol_weight=tol_weight, tol_length=tol_length, MILP_solver=MILP_solver, to_segregate=Active_Segregation*z.segregate_in_practice, to_display=False, max_trucks=max_trucks)
            #solve optimization problem 
            
        except:

            continue
            

        if a.solution!='Not ready for shipment.':
        #if we got a non-empty solution do the following
            
            residual_deliveries=pd.concat([residual_deliveries,
                a.leftout.loc[:,["Material","Description",
                "Name of sold-to party","Purch.Doc.","Delivery quantity","DPrio","DlvTy","Days Until CDD",'Deliv.date',
                'Plan GI Dt', 'Lead Times']]], ignore_index=True)
            
            residual_deliveries_stage_1=pd.concat([residual_deliveries,
                a.leftout.loc[:,["Material","Description",
                "Name of sold-to party","Purch.Doc.","Delivery quantity","DPrio","DlvTy","Days Until CDD",'Deliv.date',
                'Plan GI Dt', 'Lead Times']]], ignore_index=True)
            #add to residues leftover STO lines
            
            items_df=pd.concat([items_df,a.output], ignore_index=True)
            #stack STO lines for different locations
            df3=a.output.iloc[:,0:2].drop_duplicates()
            #get different trucks to be consolidated

            index_without_lanes=list()
            for i in a.solution.z.extract_values().keys():
                index_without_lanes.append((i[0],i[1]))
                #creating index to iterate ignoring lanes

            for k in index_without_lanes:
                df3.loc[df3.loc[:,"Truck"]==k[0]+1,k[1]]=a.solution.z.extract_values()[k[0],k[1],0]+a.solution.z.extract_values()[k[0],k[1],1]+a.solution.psi.extract_values()[k[0],k[1]]
                #get number of pallets for different pallet types

            trucks_df=pd.concat([trucks_df,df3], ignore_index=True)
            #stack trucks for different locations
            
            for j in range(0,a.choice_of_trucks):
                weights.append(a.solution.con_weight_capacity[j].body())
                lanes=a.solution.con_dummy_longitudinal_length_capacity[j].body()
                for l in range(0,2):
                    lanes=lanes+a.solution.con_dummy_length_capacity[j,l].body()
                lengths.append(lanes/2)
                #get Loading Meters per truck
                
        else:
            
            residual_deliveries=pd.concat([residual_deliveries,
                z.data.loc[:,["Material","Description",
                "Name of sold-to party","Purch.Doc.","Delivery quantity","DPrio","DlvTy","Days Until CDD", 'Deliv.date',
                'Plan GI Dt', 'Lead Times']]], ignore_index=True)
            #if we determine not to ship for a location, then add all of the STO lines considered for it
            
            residual_deliveries_stage_1=pd.concat([residual_deliveries,
                z.data.loc[:,["Material","Description",
                "Name of sold-to party","Purch.Doc.","Delivery quantity","DPrio","DlvTy","Days Until CDD", 'Deliv.date',
                'Plan GI Dt', 'Lead Times']]], ignore_index=True)
        
    if trucks_df.shape[0]>0:
        print(str((time.time()-startTime)/60) + ' minutes.')
        trucks_df.fillna(0, inplace=True)
        c=trucks_df.iloc[:,0:2].drop_duplicates()
        c['Total Pallets']=trucks_df.drop(columns=['Truck','Name of sold-to party']).apply(sum, axis=1)
        c["Weight"]=weights
        c["Loading Meters"]=lengths
        trucks_df=trucks_df.merge(c, on=['Truck','Name of sold-to party'])
        #generate trucks per location + overview of them

    trucks_df["Capacity Filled"]=trucks_df.loc[:,["Loading Meters","Weight"]].apply(lambda x: str(np.round(100*max(x[0]/13200,x[1]/23500),2))+'%',axis=1)
    #specify maximum capacity in terms of maximum percentage of filling, between loading meters and weight  
        
    for i in range(0,trucks_df.shape[0]):
        if sum(items_df.loc[(items_df.Truck==trucks_df.loc[i,"Truck"]) & (items_df.loc[:,"Name of sold-to party"]==trucks_df.loc[i,"Name of sold-to party"]),"DPrio"]==55)>0:
            trucks_df.loc[i,"Priority 55"]=True
        else:
            trucks_df.loc[i,"Priority 55"]=False
            #specify whether on a specific truck there are priority 55 orders

        if sum(items_df.loc[(items_df.Truck==trucks_df.loc[i,"Truck"]) & (items_df.loc[:,"Name of sold-to party"]==trucks_df.loc[i,"Name of sold-to party"]),"Days Until CDD"]<=0)>0:

            trucks_df.loc[i,"Due Soon"]=True
        else:
            trucks_df.loc[i,"Due Soon"]=False
            #specify whether on a specific truck there are STO lines due soon
    
    return items_df, trucks_df, residual_deliveries


# In[117]:


print("****************************************************************************************************************************")
print("Initialize Phase 1")
print("****************************************************************************************************************************")
stage_1_timeframe=7 #specify timeframe at stage 1
items_df, trucks_df, residual_deliveries=execute_optimizer(deliveries, dates, max_time=40, MILP_solver='cbc', Lambda=1.15, max_days=stage_1_timeframe, max_trucks=3)
trucks_df["Stage"]=1
dummy_trucks=items_df.copy()
#prunning of residual pallets

if items_df.shape[0]>0:

    ind=items_df.loc[(items_df.loc[:,"Quantity Delivered"]/items_df.loc[:,"Delivery quantity"]<1) & (items_df.loc[:,"Quantity Delivered"]/items_df.loc[:,"Delivery quantity"]>0.95),:].index
    #get STO lines for which we are getting more than the 95%. Then get those STO lines up to 100%
    
    if len(ind)>0:

        items_df.loc[ind,"Quantity Delivered"]=items_df.loc[ind,"Delivery quantity"]

    ss=items_df.loc[(items_df.loc[:,"Delivery quantity"]-items_df.loc[:,"Quantity Delivered"])>0,:]
    items_df.loc[ss.index,"Quantity Delivered"]=(ss.loc[:,"Delivery quantity"]/ss.loc[:,"NPallets"])*np.floor(ss.loc[:,"Quantity Delivered"]/(ss.loc[:,"Delivery quantity"]/ss.loc[:,"NPallets"]))
    ss=items_df.loc[((items_df.loc[:,"Delivery quantity"]-items_df.loc[:,"Quantity Delivered"])>0) & (items_df.loc[:,"Quantity Delivered"]>0),:]
    #get STO lines partially being served and deliver a quantity such that we deliver floor(number of pallets)
    
    residual_deliveries=pd.concat([residual_deliveries,
                    items_df.loc[((items_df.loc[:,"Delivery quantity"]-items_df.loc[:,"Quantity Delivered"])>0) & 
                    (items_df.loc[:,"Quantity Delivered"]==0),:].loc[:,["Material","Description",
                    "Name of sold-to party","Purch.Doc.","Delivery quantity","DPrio","DlvTy","Days Until CDD",'Deliv.date',
                    'Plan GI Dt', 'Lead Times']]], ignore_index=True)
    #make proper adjustments to residual deliveries
    
    residual_deliveries_stage_1=pd.concat([residual_deliveries,
                    items_df.loc[((items_df.loc[:,"Delivery quantity"]-items_df.loc[:,"Quantity Delivered"])>0) & 
                    (items_df.loc[:,"Quantity Delivered"]==0),:].loc[:,["Material","Description",
                    "Name of sold-to party","Purch.Doc.","Delivery quantity","DPrio","DlvTy","Days Until CDD",'Deliv.date',
                    'Plan GI Dt', 'Lead Times']]], ignore_index=True)
    
    items_df=items_df.loc[items_df.loc[:,"Quantity Delivered"]>0,:]
    items_df.loc[ss.index,"NPallets"]=np.floor(ss.loc[:,"Quantity Delivered"]/(ss.loc[:,"Delivery quantity"]/ss.loc[:,"NPallets"]))
    #make changes to number of pallets per STO line affected by rounding down the number of pallets
    
    ss.loc[:,"NPallets"]=ss.loc[:,"Delivery quantity"]/ss.loc[:,"NPallets"]*(ss.loc[:,"Delivery quantity"]-ss.loc[:,"Quantity Delivered"])
    ss.loc[:,"Delivery quantity"]=ss.loc[:,"Delivery quantity"]-ss.loc[:,"Quantity Delivered"]

    residual_deliveries=pd.concat([residual_deliveries,
                    ss.loc[:,["Material","Description",
                    "Name of sold-to party","Purch.Doc.","Delivery quantity",
                    "DPrio","DlvTy","Days Until CDD",'Deliv.date',
                    'Plan GI Dt', 'Lead Times']]], ignore_index=True)
    #make relevant changes to residual deliveries
    
    residual_deliveries_stage_1=pd.concat([residual_deliveries,
                    ss.loc[:,["Material","Description",
                    "Name of sold-to party","Purch.Doc.","Delivery quantity",
                    "DPrio","DlvTy","Days Until CDD",'Deliv.date',
                    'Plan GI Dt', 'Lead Times']]], ignore_index=True)

    a=residual_deliveries.loc[(residual_deliveries.loc[:,"DPrio"]==55) | (residual_deliveries.loc[:,"Days Until CDD"]<=0),"Name of sold-to party"]
    residual_deliveries=residual_deliveries.loc[np.isin(residual_deliveries.loc[:,"Name of sold-to party"],a),:]
    #get just locations for which there are priority 55 orders or there's late delivery risk to run stage 2

if residual_deliveries.shape[0]>0:

        print("")
        print("****************************************************************************************************************************")
        print("Initialize Phase 2")
        print("****************************************************************************************************************************")

        residual_items_df, residual_trucks_df, residual_deliveries2=execute_optimizer(residual_deliveries, dates, tol_weight=0, tol_length=0, max_time=30,
                                                                   MILP_solver='cbc', Lambda=1.1, max_days=21, max_trucks=1)
        #at stage 2 we re-run the algorithm just for locations with unattended priorities and just aiming to fill one truck
        #at a maximum. We consider the STO lines left out at Stage 1 and optionally consider bigger timeframes
        residual_trucks_df['Stage']=2
        items_df2=pd.concat([items_df,residual_items_df], ignore_index=True)
        trucks_df2=pd.concat([trucks_df,residual_trucks_df], ignore_index=True).fillna(0)
        
items_df.loc[:,"Delivery quantity"]=items_df.loc[:,"Delivery quantity"]-items_df.loc[:,"Quantity Delivered"]
items_df=items_df.rename(columns={"Delivery Quantity":"Split Leftover"})
residual_items_df.loc[:,"Delivery quantity"]=residual_items_df.loc[:,"Delivery quantity"]-residual_items_df.loc[:,"Quantity Delivered"]
residual_items_df=residual_items_df.rename(columns={"Delivery Quantity":"Split Leftover"})


# In[ ]:


trucks_df2


# In[132]:


directory='C:\\Users\\frand\\Documents'+'\\'+"IFFOPT_output_"+datetime.date.today().strftime('%d.%m.%Y')+'.xlsx'
timeframe_string='Left Out of Stage 1 ('+str(stage_1_timeframe)+' days)'
writer=pd.ExcelWriter(directory)
items_df.to_excel(writer,'Items Stage 1', index=False)
trucks_df2.to_excel(writer,'Trucks', index=False)
residual_items_df.to_excel(writer,'Items Stage 2', index=False)
if residual_deliveries_stage_1.shape[0]>0:
    residual_deliveries_stage_1.to_excel(writer, timeframe_string, index=False)

for column in items_df:
    column_width = max(items_df[column].astype(str).map(len).max(), len(column))
    col_idx = items_df.columns.get_loc(column)
    writer.sheets['Items Stage 1'].set_column(col_idx, col_idx, column_width)

for column in trucks_df2:
    column_width = max(trucks_df2[column].astype(str).map(len).max(), len(column))+1
    col_idx = trucks_df2.columns.get_loc(column)
    writer.sheets['Trucks'].set_column(col_idx, col_idx, column_width)

for column in residual_items_df:
    column_width = max(residual_items_df[column].astype(str).map(len).max(), len(column))
    col_idx = residual_items_df.columns.get_loc(column)
    writer.sheets['Items Stage 2'].set_column(col_idx, col_idx, column_width)

if residual_deliveries_stage_1.shape[0]>0:
    for column in residual_deliveries_stage_1:
        column_width = max(residual_deliveries_stage_1[column].astype(str).map(len).max(), len(column))
        col_idx = residual_deliveries_stage_1.columns.get_loc(column)
        writer.sheets[timeframe_string].set_column(col_idx, col_idx, column_width)
        
writer.save()
writer.close()


# In[113]:


missing_materials=deliveries["Material"][np.isin(deliveries["Material"], materials["Material"])==False].reset_index().iloc[:,1:]
missing_untits_per_pallet=new.loc[new["GR slips Essers"]==0,'Material']
if missing_materials.shape[0]>0 or missing_untits_per_pallet.shape[0]>0:
    string_new_materials="The following materials should be added to the Material Master Data in the Segregation tool:"
    string_GR_slips_error="The following materials are missing important data on the Material Master Data in the Segregation tool:"

    for i in range(0,missing_materials.shape[0]):
        string_new_materials+=(" "+str(missing_materials.iloc[i,0]))+"/"
    for i in range(0,missing_untits_per_pallet.shape[0]):
        string_GR_slips_error+=(" "+str(missing_untits_per_pallet.iloc[i,0]))+"/"
        
    #then send those strings in a email at next outlook's body. Manish please add.


# In[ ]:


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'franciscodelioperego@gmail.com' #mail to transportation planners etc
mail.Subject = 'IFFOPT Output'
mail.Body = 'No, ahí.'

# To attach a file to the email (optional):
attachment  = directory
mail.Attachments.Add(attachment)

mail.Send()

