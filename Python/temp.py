import json
import re
import pprint


def pretty_printer(obj):
        pp = pprint.PrettyPrinter(indent=4)
        pp.pprint(obj)


rex = re.compile(r'\(.*\)', re.IGNORECASE)

with open(r'C:\Users\ric\projects\wp_dot1x\Python\dot1x_sites_dates.json', 'r') as f:
    data = json.load(f)

with open(r'C:\Users\ric\projects\wp_dot1x\Python\device_info.json', 'r') as fin:
    device_info = json.load(fin)

new_dict = {}
for k, rest in data.items():
    text_desc = rex.sub('', rest['switches'])
    text_desc = text_desc.strip()
    split_at = re.compile(r'(?:,\s|\sand\s)')
    for switch in split_at.split(text_desc):
        if k not in new_dict:
                new_dict[k] = []
        new_dict[k].append(switch)
arr = []
last_location = ''
for location, switches in new_dict.items():
    if location != last_location:
        arr.append([location, '', ''])# , '', ''])
    for switch in switches:
        if switch in device_info:
            model = device_info[switch]['device_model']
            device_serial = device_info[switch]['device_serial']
            version = device_info[switch]['version']
            arr.append([location, switch, model]) #, device_serial, version])
        elif switch[:-2] in device_info: #Switch is part of stack
            model = f"{device_info[switch[:-2]]['device_model']} (part of stack)"
            device_serial = 'TBA'
            version = device_info[switch[:-2]]['version']
            arr.append([location, switch, model]) #, device_serial, version])
        else:
            print(f'Switch {switch} not found in device info dump!')
    last_location = location

arr = sorted(arr, key=lambda x: [x[0], x[1]])

for entry in arr:
    print("|".join(entry))