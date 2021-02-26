# -*- coding: utf-8 -*-
"""
Created on Sat Jan 30 12:36:24 2021

@author: JLuca
"""
# -*- coding: utf-8 -*-
"""
Created on Sat Jan 23 17:22:29 2021

@author: JLuca
"""

import pandas as pd
from datetime import date
from forex_python.converter import CurrencyRates

def Trending_Value(data: pd.DataFrame):
    ''' Function computes the Trending Value portfolio and outputs it to an Excel file. 
        Trending Value is computed as first selecting the 10% cheapest stocks, determined by a composite value measure.
        The most undervalued stocks are then ranked by strongest combined momentum to get the final ranking for stock selection.
    '''
    
    # Select relevant columns
    DATA = data[['ID', 'Company name', 'Ticker', 'Country', 'Industry',
                 'Date', 'Share Price', 'EV', 'P/E','P/S', 'P/B',
                 'EBIT/EV', 'P/FCF', 'Yield','PCHG 3m', 'PCHG 6m','PCHG 12m']]
        
    # Fill NaNs for some columns
    DATA['Yield'] = DATA['Yield'].fillna(0)
    
    
    # Drop rows with NaN values
    DATA = DATA.dropna()
    
    # Convert Enterprise Value to SEK
    DATA = _convert_to_SEK(DATA)    
       
    # 2. Filter size
    minSize = 250 # MSEK
    DATA = DATA[ DATA['EV'] >= minSize ]
        
    
    # Create metrics for Trending Value
    DATA['E/P'] = 1/DATA['P/E']
    DATA['S/P'] = 1/DATA['P/S']
    DATA['B/P'] = 1/DATA['P/B']
    DATA['FCF/P'] = 1/DATA['P/FCF']
    DATA['Combined Momentum'] =  (DATA['PCHG 3m'] 
                                   + DATA['PCHG 6m'] 
                                   + DATA['PCHG 12m'])/3
    
    
    # Create rankings
    # ascending=False --> Highest value is given Rank = 1
    DATA['E/P Rank'] = DATA['E/P'].rank(ascending=False) 
    DATA['S/P Rank'] = DATA['S/P'].rank(ascending=False)
    DATA['B/P Rank'] = DATA['B/P'].rank(ascending=False)
    DATA['EBIT/EV Rank'] = DATA['EBIT/EV'].rank(ascending=False)
    DATA['FCP/P Rank'] = DATA['FCF/P'].rank(ascending=False)
    DATA['Yield Rank'] = DATA['Yield'].rank(ascending=False)
    
    
    # Sum value points
    DATA['Combined Value Rank Points'] = DATA['E/P Rank'] + DATA['S/P Rank'] + DATA['B/P Rank'] \
                                       + DATA['EBIT/EV Rank'] + DATA['FCP/P Rank'] + DATA['Yield Rank']
                            
    
    # Rank by Value and keep the 10% best valued stocks
    DATA['Combined Value Rank'] = DATA['Combined Value Rank Points'].rank(ascending=True)
    DATA = DATA.sort_values(by=['Combined Value Rank'])
    Tenth_percentile = int(0.1*DATA.shape[0])
    DATA = DATA[ DATA['Combined Value Rank']<= Tenth_percentile ]
    
    
    # Calculate Momentum ranking for the 10% best valued stocks
    DATA['Combined Momentum Rank'] = DATA['Combined Momentum'].rank(ascending=False)
    
    
    # Sort by momentum to get final ranking for stock selection
    DATA = DATA.sort_values(by=['Combined Momentum Rank'])
        
    
    # Save to Excel
    DATA.to_excel(f"1_Trending_Value_{date.today()}.xlsx") 
        
def Trending_Dividend(data: pd.DataFrame):
    ''' Function computes the Trending Dividend portfolio and outputs it to an Excel file. 
        Trending Dividend is computed as first selecting the 10% stocks with the highest dividend.
        The highest dividend stocks are then ranked by strongest combined momentum to get the final ranking for stock selection.
    '''
    
    # Select relevant columns
    DATA = data[['ID', 'Company name', 'Ticker', 'Country', 'Industry', 'Date',
                 'Share Price', 'EV', 'Yield','PCHG 3m', 'PCHG 6m','PCHG 12m']]
        
    
    # Fill NaNs for some columns
    DATA['Yield'] = DATA['Yield'].fillna(0)
    
    
    # Drop rows with NaN values
    DATA = DATA.dropna()

    
    # Convert Enterprise Value to SEK
    DATA = _convert_to_SEK(DATA)   
     
       
    # Filter size
    minSize = 250 # MSEK
    DATA = DATA[ DATA['EV'] >= minSize ]
        
    
    # Create metrics for Trending Dividend
    DATA['Combined Momentum'] =  (DATA['PCHG 3m'] 
                                   + DATA['PCHG 6m'] 
                                   + DATA['PCHG 12m'])/3
    
    
    # Create rankings
    # ascending=False --> Highest value is given Rank = 1
    DATA['Yield Rank'] = DATA['Yield'].rank(ascending=False)
    
    
    # Rank by Yield and keep the 10% best valued stocks
    Tenth_percentile = int(0.1*DATA.shape[0])
    DATA = DATA[ DATA['Yield Rank']<= Tenth_percentile ]
    
    
    # Calculate Momentum ranking for the 10% best valued stocks
    DATA['Combined Momentum Rank'] = DATA['Combined Momentum'].rank(ascending=False)
    
    
    # Sort by momentum to get final ranking for stock selection
    DATA = DATA.sort_values(by=['Combined Momentum Rank'])
        
    
    # Save to Excel
    DATA.to_excel(f"2_Trending_Dividend_{date.today()}.xlsx") 
     
def Trending_Quality(data: pd.DataFrame):
    ''' Function computes the Trending Quality portfolio and outputs it to an Excel file. 
        Trending Quality is computed as first selecting the 10% stocks with the highest quality, determined by a composite measure.
        The highest quality stocks are then ranked by strongest combined momentum to get the final ranking for stock selection.
    '''
    
    # Select relevant columns
    DATA = data[['ID', 'Company name', 'Ticker', 'Country', 'Industry', 'Date',
                 'Share Price', 'EV', 'Equity', 'FCF', 'ROIC', 'ROA', 'ROE', 
                 'PCHG 3m', 'PCHG 6m','PCHG 12m']]
    
        
    # Drop rows with NaN values
    DATA = DATA.dropna()
    

    # Convert Enterprise Value to SEK
    DATA = _convert_to_SEK(DATA)   
            
    # 2. Filter by size
    minSize = 250 # MSEK
    DATA = DATA[ DATA['EV'] >= minSize ]
        
    
    # Create metrics for Trending Quality
    DATA['FCF/Equity'] = DATA['FCF'] / DATA['Equity']
    DATA['Combined Momentum'] =  (DATA['PCHG 3m'] 
                                   + DATA['PCHG 6m'] 
                                   + DATA['PCHG 12m'])/3
    

    # Create rankings
    # ascending=False --> Highest value is given Rank = 1
    DATA['FCF/Equity Rank'] = DATA['FCF/Equity'].rank(ascending=False) 
    DATA['ROIC Rank'] = DATA['ROIC'].rank(ascending=False)
    DATA['ROA Rank'] = DATA['ROA'].rank(ascending=False)
    DATA['ROE Rank'] = DATA['ROE'].rank(ascending=False)
    
    
    # Sum value points
    DATA['Combined Value Rank Points'] = DATA['FCF/Equity Rank'] + DATA['ROIC Rank'] \
                                        + DATA['ROA Rank'] + DATA['ROE Rank']
                            
    
    # Rank by Value and keep the 10% best valued stocks
    DATA['Combined Value Rank'] = DATA['Combined Value Rank Points'].rank(ascending=True)
    DATA = DATA.sort_values(by=['Combined Value Rank'])
    Tenth_percentile = int(0.1*DATA.shape[0])
    DATA = DATA[ DATA['Combined Value Rank']<= Tenth_percentile ]
    
    
    # Calculate Momentum ranking for the 10% best valued stocks
    DATA['Combined Momentum Rank'] = DATA['Combined Momentum'].rank(ascending=False)
    
    
    # Sort by momentum to get final ranking for stock selection
    DATA = DATA.sort_values(by=['Combined Momentum Rank'])
        
    
    # Save to Excel
    DATA.to_excel(f"3_Trending_Quality_{date.today()}.xlsx") 

def Combined_Momentum(data: pd.DataFrame):
    ''' Function computes the Combined Momentum portfolio and outputs it to an Excel file. 
        Combined momentum is computed by first filtering out companies with Piotroski F-score lower than 6.
        The resulting stocks are then ranked by strongest combined momentum to get the final ranking for stock selection.
    '''
    
    # Select relevant columns
    DATA = data[['ID', 'Company name', 'Ticker', 'Country', 'Industry', 'Date',
                 'Share Price', 'EV', 'F-Score', 'PCHG 3m', 'PCHG 6m','PCHG 12m']]
    
        
    # Drop rows with NaN values
    DATA = DATA.dropna()
    
    
    # Convert Enterprise Value to SEK
    DATA = _convert_to_SEK(DATA)   
    
        
    # Filter size
    minSize = 250 # MSEK
    DATA = DATA[ DATA['EV'] >= minSize ]
    
    
    # Filter out low F-Score
    minF = 6
    DATA = DATA[ DATA['F-Score'] >= minF  ]
        
    
    # Create metrics for Trending Dividend
    DATA['Combined Momentum'] =  (DATA['PCHG 3m'] 
                                   + DATA['PCHG 6m'] 
                                   + DATA['PCHG 12m'])/3
        
    
    # Calculate Momentum ranking
    DATA['Combined Momentum Rank'] = DATA['Combined Momentum'].rank(ascending=False)
    
    
    # Sort by momentum to get final ranking for stock selection
    DATA = DATA.sort_values(by=['Combined Momentum Rank'])
        
    
    # Save to Excel
    DATA.to_excel(f"4_Combined_Momentum_{date.today()}.xlsx") 

def Net_Net(data: pd.DataFrame, minF = 4):
    ''' Function computes the Net Net portfolio and outputs it to an Excel file. 

    '''
    
    # Net Current Asset Value (NCAV) := Current Assets - Total liabilites (incl. preferred stock)
    
    # Select relevant columns
    DATA = data[['ID', 'Company name', 'Ticker', 'Country', 'Industry',
                 'Date', 'F-Score', 'Share Price', 'EV', 'MV', 'Current assets', 'Total debt']]
        
    # Drop rows with NaN values
    DATA = DATA.dropna()
    
    # Filter F-score
    DATA = DATA[ DATA['F-Score'] >= minF ]
        
    # Create metrics for NCAV
    DATA['NCAV'] = DATA['Current assets'] - DATA['Total debt']  
    
    # Filter out NCAV < 0
    DATA = DATA[ DATA['NCAV']>0 ]

    # Create metric for NCAV/MV
    DATA['MV/NCAV'] =  DATA['MV'] / DATA['NCAV']
    
    # Create ranking
    # ascending = True --> Lowest value is given Rank = 1
    DATA['MV/NCAV Rank'] = DATA['MV/NCAV'].rank(ascending=True) 
        
    # Sort by ranking
    DATA = DATA.sort_values(by=['MV/NCAV Rank'])
        
    # Save to Excel
    DATA.to_excel(f"5_Net_Net_{date.today()}.xlsx") 

def Deep_Value(data: pd.DataFrame, value_param = 'EBIT'): 
    ''' Function computes the Deep Value portfolio and outputs it to an Excel file. 

    '''
    
    # Deep Value: Cheapest stocks determined as EV/EBIT
    
    # Select relevant columns
    DATA = data[['ID', 'Company name', 'Ticker', 'Country', 'Industry',
                 'Date', 'EV', 'Cash', value_param]]
        
    # Drop rows with NaN values
    DATA = DATA.dropna()
    
    # Filter out negative value param
    DATA = DATA[ DATA[value_param]>=0 ]
    
    # Create metrics
    DATA['TEV'] = DATA['EV'] - DATA['Cash']
    
    DATA[f'TEV/{value_param}'] = DATA['TEV']/DATA[value_param]
    
    # Create ranking
    # ascending = True --> Lowest value is given Rank = 1
    DATA[f'TEV/{value_param} Rank'] = DATA[f'TEV/{value_param}'].rank(ascending=True) 
        
    # Sort by ranking
    DATA = DATA.sort_values(by=[f'TEV/{value_param}'])
        
    # Save to Excel
    DATA.to_excel(f"6_Deep_Value_{date.today()}.xlsx") 
    
def _convert_to_SEK(DATA):
    ''' Function returns the DataFrame with the Enterprise Value (EV) column values changed from EURO, DKK and USD to SEK '''
    Finnish_companies = DATA.loc[DATA['Country']=='Finland', 'Company name']
    Danish_companies = DATA.loc[DATA['Country']=='Danmark', 'Company name']
    USA_companies = DATA.loc[DATA['Country']=='USA', 'Company name']
    for i, row in DATA.iterrows():
        if i in Finnish_companies.index:
            DATA.at[i,'EV'] *= exchange_rates['EUR_SEK']
        elif i in Danish_companies.index:
            DATA.at[i,'EV'] *= exchange_rates['DKK_SEK']
        elif i in USA_companies.index:
            DATA.at[i,'EV'] *= exchange_rates['USD_SEK']
    return DATA

def change_column_names(DATA):
    ''' Function returns the input DataFrame but with changed column names'''
    
    return DATA.rename(columns = {
               'Börsdata ID' : 'ID', 'Bolagsnamn' : 'Company name', 'Info - Ticker': 'Ticker', 'Info - Land': 'Country',
               'Info - Bransch' : 'Industry', 'Info - Aktiekurs' : 'Date' ,'Aktiekurs - Senaste': 'Share Price', 'F-Score - Poäng' : 'F-Score',
               'Kassa - Miljoner' : 'Cash', 'FCF - Miljoner': 'FCF', 'OP kassaf. - Miljoner': 'OCF', 'EBIT - Miljoner' : 'EBIT', 
               'Börsvärde - Senaste' : "MV",'EV - Senaste': 'EV', 'EBIT/EV (%) - Senaste' : 'EBIT/EV', 'Omsätt.tillg. - Miljoner' : 'Current assets',
               'Tot. Skulder - Miljoner' : 'Total debt', 'Eget Kapital - Miljoner' : 'Equity', 'ROA - Senaste' : 'ROA',
               'ROE - Senaste' : 'ROE', 'ROC - Senaste' : 'ROC', 'ROIC - Senaste' :'ROIC', 'P/E - Senaste': 'P/E',
               'P/S - Senaste' : 'P/S', 'P/B - Senaste' : 'P/B', 'EBIT/EV (%) - Senaste' : 'EBIT/EV',
               'P/FCF - Senaste' : 'P/FCF', 'Direktav. - Senaste' : 'Yield', 'Kursutveck. - Utveck.  3m' :'PCHG 3m',
               'Kursutveck. - Utveck.  6m' : 'PCHG 6m', 'Kursutveck. - Utveck.  1 år' : 'PCHG 12m'
             })
            
# Get exchange rates for Enterprise Value conversion to SEK
global exchange_rates
exchange_rates = {'EUR_SEK': CurrencyRates().get_rate('EUR', 'SEK'),
                  'DKK_SEK': CurrencyRates().get_rate('DKK', 'SEK'),
                  'USD_SEK': CurrencyRates().get_rate('USD', 'SEK')}
 
# Specify input file name 
file_name = f"Borsdata_{date.today()}.xlsx"

# Load input data
DATA = pd.read_excel(file_name, sheet_name='Export')

# Change columns names on the input data
DATA = change_column_names(DATA)

# Construct stock portfolios
Trending_Value(data = DATA) # Outputs Excel
Trending_Dividend(data = DATA) # Outputs Excel file
Trending_Quality(data = DATA) # Outputs Excel file
Combined_Momentum(data = DATA) # Outputs Excel file
Net_Net(data = DATA, minF = 0) # Outputs Excel file
Deep_Value(data = DATA, value_param ='EBIT') # Outputs Excel file

