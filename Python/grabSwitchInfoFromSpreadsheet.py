import win32com.client
import pprint
import json


def pretty_printer(obj):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(obj)


class Excel:
    def __init__(self, path):
        self.path = path
        self.xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
        self.xlApp.Visible = True
        self.wk = self.xlApp.Workbooks.Open(self.path)
        self.sh = self.wk.Worksheets('Data')

    @property
    def getEOL(self):
        eof = self.sh.Range('B65536').End(-4162).Row
        return eof

    @property
    def getData(self):
        data = {}
        for row in range(2, self.getEOL+1):
            hostname = self.sh.Range(f'A{row}').Value if self.sh.Range(f'A{row}').Value else None
            if hostname and hostname.lower() != 'total':
                hostname = hostname.upper()
                device_model = self.sh.Range(f'E{row}').Value
                device_serial = self.sh.Range(f'F{row}').Value
                version = self.sh.Range(f'C{row}').Value
                entry = dict(
                             device_model=device_model,
                             device_serial=device_serial,
                             version=version
                             )
                if hostname not in data:
                    data[hostname] = {}
                data[hostname] = entry
        return data

    def to_json(self, file_path):
        with open(file_path, 'w') as f:
            json.dump(self.getData, f, indent=4)


if __name__ == '__main__':
    xl = Excel(r'C:\Users\ric\Desktop\Software Currency Report Data.xlsx')
    print(xl.getEOL)
    pretty_printer(xl.getData)
    xl.to_json(file_path=r'C:\Users\ric\projects\wp_dot1x\Python\device_info.json')
