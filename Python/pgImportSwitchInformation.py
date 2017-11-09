import re
import psycopg2
import json
import pprint


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


with open(r'C:\Users\rdapaz\projects\wp_dot1x\Python\device_info.json', 'r') as fin:
    data = json.load(fin)

arr = []
for switch, p in data.items():
    arr.append([switch, p['device_model'], p['device_serial'], p['version']])

with psycopg2.connect("dbname='EOL' user=postgres") as conn:
    try:
        cur = conn.cursor()

        sql = """
            CREATE TABLE IF NOT EXISTS "Switches" (
               switch_name TEXT,
               device_model TEXT,
               device_serial TEXT,
               version TEXT
               )
               """

        cur.execute(sql)

        sql =  """ 
                INSERT INTO \"public\".\"Switches\" (switch_name, device_model, device_serial, version) 
                VALUES (%s, %s, %s, %s)
                """
        cur.executemany(sql, arr)
        conn.commit()
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
