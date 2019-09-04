#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# by Igor Pimenta - https://github.com/igorjpimenta

from blip_report_requisitor import Requisitor
from datetime import datetime
from datetime import timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Google Drive API
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# Google Sheets credentials
credentials = ServiceAccountCredentials.from_json_keyfile_name('BLiP to G-Sheets.json', scope)
gc = gspread.authorize(credentials)

# Open workbook
wks = gc.open('workbook name')

# Clear data by resize
wks.sheet1.resize(rows = 1)
    
# BOT Builder key list
key_list = ['1st key',
            '2nd key',
            '3rd key',
            '[...] key']
# BOT Builder name list
key_name = ['1st name',
            '2nd name',
            '3rd name',
            '[...] name']

# Create rows list to input on sheet
rows = []
ref = 0

# Insert data header
header = ['buider','category','action','amount','date']
rows.append(header)
wks.values_update('Cell in A1 notation ex: Plan1!A1',
                  params={'valueInputOption': 'RAW'},
                  body={'values': rows})
    
for key in key_list:    
    ref = ref + len(rows)
    
    # Clear rows list
    rows = []
    
    # Data date range
    for day in range(1,31):
        start_date = datetime.now() - timedelta(days=day)
        
        # Set BOT Buider key
        req = Requisitor(key)
        
        # Create/clear event track list
        t = []
        
        # Trackings category list
        trackings = ['1st tracking',
                     '2nd tracking',
                     '3rd tracking',
                     '[...] tracking']
    
        for i in trackings:
            # HTML decode caracters
            rep = i.replace(' ', '%20')
            rep = rep.replace('à','%C3%A0')
            rep = rep.replace('À','%C3%80')
            rep = rep.replace('á','%C3%A1')
            rep = rep.replace('Á','%C3%81')
            rep = rep.replace('â','%C3%A2')
            rep = rep.replace('Â','%C3%82')
            rep = rep.replace('ã','%C3%A3')
            rep = rep.replace('ó','%C3%B3')
            rep = rep.replace('Ó','%C3%93')
            rep = rep.replace('é','%C3%A9')
            rep = rep.replace('É','%C3%89')
            rep = rep.replace('ê','%C3%AA')
            rep = rep.replace('ú','%C3%BA')
            rep = rep.replace('Ú','%C3%9A')
            rep = rep.replace('Ê','%C3%8A')
            rep = rep.replace('í','%C3%AD')
            rep = rep.replace('Í','%C3%8D')
            rep = rep.replace('õ','%C3%B5')
            rep = rep.replace('ç','%C3%A7')
            rep = rep.replace('/','%2F')
            t.append('/event-track/' + rep)
    
        reports = [req.getCustomReport(x, start_date, start_date) for x in t]
        
        start_date = start_date.strftime('%d/%m/%Y')
        for count, category in enumerate(reports):
            if type(category) is not int:
                for action in category:
                        row = [key_name[key_list.index(key)],trackings[count],action['acao'],action['total'],start_date]
                        rows.append(row)                       
    
    wks.sheet1.resize(rows = ref + len(rows) + 1)
    wks.values_update('Cell in A1 notation, except number ex: Plan1!A' + str(ref + 1),
                      params={'valueInputOption': 'RAW'},
                      body={'values': rows})

print('Finish')
