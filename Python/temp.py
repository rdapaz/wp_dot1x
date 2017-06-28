import json

with open(r'C:\Users\rdapaz\projects\wp_dot1x\Python\dot1x_sites_dates.json', 'r') as f:
    data = json.load(f)

for k, _ in data.items():
    print(k)

sites = {
        'FORR  ( 3 X VG and 1x AP)': 'T1 - Forrestfield',
        'HO (6 x VG Devices and 2 x RB)': 'T1 - HO',
        'KALG (1 device)': 'T4 - Kalg',
        'MARGR (1 device)': 'T4 - Mriver',
        'NTHM (2 devices)': 'T2 - Northam',
        'PICT (2 devices)': 'T1 - Picton',
        'WRNA(1 x Router and 1 x RB)': 'T4 - Waroona',
        'MTCL (2 x AP)': 'T1 - Mt.Claremont',
        'STRLG (8 x AP)': 'T1 - Stirling',
        'GERN (3 x AP)': 'T2 - Geraldton',
        'Jerra (1 x RB)': 'T4 - Jerramungup',
        'Colli (1 x RB)': 'T4 - Collie',
}