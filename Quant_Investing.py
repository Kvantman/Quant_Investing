# -*- coding: utf-8 -*-
"""
Created on Sat Jan 23 17:22:29 2021

@author: JLuca
"""

import pandas as pd
from datetime import date


def Trending_Value(file_name: str):
    
    # Load Data
    DATA = pd.read_excel(file_name, sheet_name='Export')
    
    
    # Rename columns
    DATA.columns = ['Börsdata ID', 'Bolagsnamn', 'Ticker', 'Land', 'Bransch',
                    'Info - Aktiekurs', 'Aktiekurs', 'EV',
                    'P/E', 'P/S', 'P/B',
                    'EBIT/EV', 'P/FCF', 'Yield',
                    'PCHG 3m', 'PCHG 6m',
                    'PCHG 12m']
    
    
    # Fill NaNs for some columns
    DATA['Yield'] = DATA['Yield'].fillna(0)
    
    
    # Drop rows with NaN values
    DATA = DATA.dropna()
    
    # Convert Enterprise Value to SEK
    DATA = _Convert_to_SEK(DATA)    
       
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
        
def Trending_Dividend(file_name: str):
    
    # Load Data
    DATA = pd.read_excel(file_name, sheet_name='Export')
    
    
    # Rename columns
    DATA.columns = ['Börsdata ID', 'Bolagsnamn', 'Land', 'Bransch',
                    'Info - Aktiekurs', 'Aktiekurs', 'EV', 'Yield',
                    'PCHG 3m', 'PCHG 6m','PCHG 12m']
    
    
    # Fill NaNs for some columns
    DATA['Yield'] = DATA['Yield'].fillna(0)
    
    
    # Drop rows with NaN values
    DATA = DATA.dropna()

    
    # Convert Enterprise Value to SEK
    DATA = _Convert_to_SEK(DATA)   
     
       
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
     
def Trending_Quality(file_name: str):
    
    # Load Data
    DATA = pd.read_excel(file_name, sheet_name='Export')
    
    
    # Rename columns
    DATA.columns = ['Börsdata ID', 'Bolagsnamn', 'Land', 'Bransch',
                    'Info - Aktiekurs', 'Aktiekurs', 'EV', 'PCHG 3m',
                    'PCHG 6m', 'PCHG 12m', 'Equity - MSEK', 'FCF - MSEK',
                    'ROIC', 'ROA', 'ROE']    
    
    # Drop rows with NaN values
    DATA = DATA.dropna()
    

    # Convert Enterprise Value to SEK
    DATA = _Convert_to_SEK(DATA)   
            
    # 2. Filter by size
    minSize = 250 # MSEK
    DATA = DATA[ DATA['EV'] >= minSize ]
        
    
    # Create metrics for Trending Quality
    DATA['FCF/Equity'] = DATA['FCF - MSEK'] / DATA['Equity - MSEK']
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

def Combined_Momentum(file_name: str):
    
    # Load Data
    DATA = pd.read_excel(file_name, sheet_name='Export')
    
    
    # Rename columns
    DATA.columns = ['Börsdata ID', 'Bolagsnamn', 'Land', 'Bransch',
                    'Info - Aktiekurs', 'Aktiekurs', 'EV', 'F-Score',
                    'PCHG 3m', 'PCHG 6m','PCHG 12m']
    

    # Drop rows with NaN values
    DATA = DATA.dropna()
    
    
    # Convert Enterprise Value to SEK
    DATA = _Convert_to_SEK(DATA)   
    
        
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

def _Convert_to_SEK(DATA):
    
    #exchange_rates = [10.17, 1.37]
    EUR_SEK = exchange_rates[0]
    DKK_SEK = exchange_rates[1]
    Finnish_companies = DATA.loc[DATA['Land']=='Finland', 'Bolagsnamn']
    Danish_companies = DATA.loc[DATA['Land']=='Danmark', 'Bolagsnamn']   
    for i, row in DATA.iterrows():
        if i in Finnish_companies.index:
            DATA.at[i,'EV'] *= EUR_SEK
        elif i in Danish_companies.index:
            DATA.at[i,'EV'] *= DKK_SEK
    return DATA


# Excel file inputs
TV_input = f"Borsdata_TV_{date.today()}.xlsx"
TD_input = f"Borsdata_TD_{date.today()}.xlsx"
TQ_input = f"Borsdata_TQ_{date.today()}.xlsx"
CM_input = f"Borsdata_CM_{date.today()}.xlsx"


# Exchange rates [EUR_SEK, DKK_SEK]
global exchange_rates
exchange_rates = [10.17, 1.37]


# Compute Trending Value Portfolio
Trending_Value(file_name = TV_input) # Outputs Excel file

# Construct Trending Dividend Portfolio
Trending_Dividend(file_name = TD_input) # Outputs Excel file

# Construct Trending Quality Portfolio
Trending_Quality(file_name = TQ_input) # Outputs Excel file

# Combined_Momentum_Portfolio
Combined_Momentum(file_name = CM_input) # Outputs Excel file