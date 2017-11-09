import re
import psycopg2
import json
import pprint


def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


with open(r'C:\Users\rdapaz\projects\wp_dot1x\Python\dot1x_sites_dates.json', 'r') as fin:
    data = json.load(fin)

pretty_printer(data)

arr = []
for tier_site, p in data.items():
    tier, site = tier_site.split(' - ', 2)
    print(tier, site, sep="|")
    tier = tier.strip()
    site = site.strip()
    arr.append([tier, site, p['switches'], p['qty']])

with psycopg2.connect("dbname='EOL' user=postgres") as conn:
    try:
        cur = conn.cursor()

        sql = """
            CREATE TABLE IF NOT EXISTS "TierSites" (
               tier TEXT,
               site TEXT,
               switches TEXT,
               qty int
               )
               """

        cur.execute(sql)

        sql =  """ 
                INSERT INTO \"public\".\"TierSites\" (tier, site, switches, qty) 
                VALUES (%s, %s, %s, %s)
                """
        cur.executemany(sql, arr)
        conn.commit()
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)