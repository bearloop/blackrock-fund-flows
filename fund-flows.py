import pandas as pd
import requests as req
import numpy as np

# Read Blackrock's products list
xl = pd.read_excel('assets/product-screener.xlsx')
cols = ['URL','Asset Class','Sub Asset Class','Region','Market']

# Filter out non UCITS ETFs
xl = xl[xl.Name.str.lower().str.contains('ucits').replace(np.NaN,False)][cols]

# Add ajax file suffix to generate the request link
xl['Download'] = xl.URL +'/1541728496966.ajax?fileType=xls'

# Fund name
xl['name'] = [i[-1] for i in xl.URL.str.split('/')]

# Fund id
xl['fund_id'] = [i[-2] for i in xl.URL.str.split('/')]



def fetch_blackrock_xml(url):
    
    # Read url
    r = req.get(url)
    
    # Specify xpath and namespace, read as xml
    res = pd.read_xml(r.text.replace('\ufeff\ufeff',''),
                xpath='''//ss:Worksheet[@ss:Name='Historical']//ss:Data''',
                namespaces={"ss": "urn:schemas-microsoft-com:office:spreadsheet"})

    # To pandas
    res = pd.DataFrame(res['Data'].values.reshape(-1,7))

    # Handle columns
    res.columns = res.loc[0].values
    res.drop(0,inplace=True)

    # Handle index/dates
    res.index = pd.to_datetime(res['As Of'])
    res.drop(['As Of'],axis=1, inplace=True)
    res.sort_index(inplace=True)

    # Replace certain values
    res.replace(['--','-',''],np.NaN, inplace=True)

    # Values to float
    res.iloc[:,1:] = res.iloc[:,1:].astype(float)

    # Set as null every row with a missing NAV
    res[res['NAV per Share'].isnull()] = np.NaN

    # Fill forward
    res.ffill(inplace=True)

    # "Flows" = daily difference in Net Assets
    res['Flows'] = res['Total Net Assets'].diff()

    # Flows as % of assets
    res['1wF/A'] = (res['Flows'].rolling(5).sum()/res['Total Net Assets'])
    res['1mF/A'] = (res['Flows'].rolling(21).sum()/res['Total Net Assets'])
    res['3mF/A'] = (res['Flows'].rolling(21*3).sum()/res['Total Net Assets'])
    res['12mF/A'] = (res['Flows'].rolling(21*12).sum()/res['Total Net Assets'])
    
    return res

# Fetch data
flows = dict()
for url in xl['Download']:
    fund_id = url.split('products/')[1].split('/')[0]
    try:
        flows[fund_id] = fetch_blackrock_xml(url)
    except:
        print('Could not retrieve data for', fund_id)
        flows[fund_id] = 'N/A'