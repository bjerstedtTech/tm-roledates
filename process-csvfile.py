#!/usr/bin/env python3
import csv
import re
from datetime import date

cutoff = date.fromisoformat('2019-09-04')
# cutoff = date.fromisoformat('2019-09-04')
roles = ['TM', 'SP', 'TT', 'GE', 'E', 'GR', 'IT', 'CJ', '1M']

with open('TechMasters Schedule.csv') as f, open('roledates.csv', 'w') as wf:
    reader = csv.DictReader(f)

    writer = csv.DictWriter(wf, ["Name"] + roles)
    writer.writeheader()

    dates = list(filter(lambda d: date.fromisoformat(d) <= cutoff, reader.fieldnames[1:]))
    
    for row in reader:
        name = row[reader.fieldnames[0]]
        
        if name == "Open Roles":
            break

        # Extract role for every date in slected range of dates -- ignoe N/A or blank
        # rd = filter(lambda p: p[0] and not re.match(r'(?i)N/A', p[0]), [(row[d], d) for d in dates])
        rd = filter(lambda p: p[0] and not re.match(r'(?i)N/A', p[0]), [(row[d], d) for d in dates])        
        
        # Split or list of multiple roles
        # rd = [(list(filter(lambda r: re.match(r'\w+', r), re.split(r'\W+', p[0]))), p[1]) for p in rd]
        rd = [(list(filter(lambda r: re.match(r'\w+', r), re.split(r'\W+', p[0]))), p[1]) for p in rd]
        
        dd = dict(filter(lambda rp: rp[0] in roles, [(r, p[1]) for p in rd for r in p[0]]))

        roledates = { "Name": name }
        roledates.update()
        for role in roles:
            roledates[role] = dd[role] if role in dd else None

        writer.writerow(roledates)
