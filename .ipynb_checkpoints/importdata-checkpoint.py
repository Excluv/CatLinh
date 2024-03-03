import pandas as pd
import sqlite3
import dateutil.parser as parser


def parsetype(type_ids):
    for x in range(len(type_ids)):
        if type_ids[x] == '40RF':
            type_ids[x] = int(1)

    return type_ids


def parsecustomer(customer_ids):
    for x in range(len(customer_ids)):
        if customer_ids[x] == 'LDP':
            customer_ids[x] = int(1)

    return customer_ids


def parsedestination(destination_ids):
    cur.execute('SELECT ID, Name FROM Destination')
    destinations = list()
    for row in cur:
        destinations.append(row[1].upper())

    for x in range(len(destination_ids)):
        if destination_ids[x] in destinations:
            destination_ids[x] = int(destinations.index(destination_ids[x])) + 1

    return destination_ids


def parselocation(type_ids, l_ids, action):
    if action == 'Pick Up':
        cur.execute('SELECT Name, Type_ID FROM PickupLocation')
    if action == 'Drop Off':
        cur.execute('SELECT Name, Type_ID FROM DropoffLocation')
    loc = cur.fetchall()
    for x in range(len(l_ids)):
        if (l_ids[x], type_ids[x]) in loc:
            l_ids[x] = int(loc.index((l_ids[x], type_ids[x]))) + 1

    return l_ids


def parsedate(dates):
    for x in range(len(dates)):
        if (dates[x] == 'nan') or (dates[x] == 'NaN') or (dates[x] == 'NAN')\
            or (dates[x] == 'Nan'): continue
        if dates[x] == None: continue
        if type(dates[x]) == float: continue
        dates[x] = str(parser.parse(dates[x], dayfirst=True)).split()
        dates[x] = dates[x][0]

    return dates

def importdata():
    df = pd.read_excel('C:\\Users\\jio\\Desktop\\main.xlsx')

    # Resolve parsing data formats:
    type_ids = parsetype(list(df.iloc[:, 0]))
    customer_ids = parsecustomer(list(df.iloc[:, 2]))
    destination_ids = parsedestination(list(df.iloc[:, 3]))
    PUL_ids = parselocation(type_ids, list(df.iloc[:, 5]), 'Pick Up')
    DOL_ids = parselocation(type_ids, list(df.iloc[:, 6]), 'Drop Off')
    from_dates = parsedate(list(df.iloc[:, 7]))
    to_dates = parsedate(list(df.iloc[:, 8]))

    # Extract raw data:
    cont_ids = list(df.iloc[:, 1])
    vehicles = list(df.iloc[:, 4])
    provisions = list(df.iloc[:, 9])
    incurred_costs = list(df.iloc[:, 11])
    descriptions = list(df.iloc[:, 12])

    print(len(type_ids), len(cont_ids), len(customer_ids),
          len(destination_ids), len(vehicles), len(PUL_ids),
          len(DOL_ids), len(from_dates), len(to_dates),
          len(provisions), len(incurred_costs), len(descriptions))

    data = list(0 for x in range(len(cont_ids)))
    for x in range(len(cont_ids)):
        try:
            data[x] = ((type_ids[x], cont_ids[x], customer_ids[x],
                   destination_ids[x], vehicles[x], PUL_ids[x],
                   DOL_ids[x], from_dates[x], to_dates[x],
                   provisions[x], incurred_costs[x], descriptions[x]))
        except:
            print('Error: ', Exception)

    # Insert data into 'Orders' table:
    command = '''
    INSERT INTO Orders (Type_ID, Container_ID, Customer_ID,
                        Destination_ID, Vehicle, PUL_ID,
                        DOL_ID, From_Date, To_Date,
                        Provision, Incurred_Cost, Description)
    VALUES (?, ?, ?,
            ?, ?, ?,
            ?, ?, ?,
            ?, ?, ?)'''

    cur.executemany(command, data)
    conn.commit()


conn = sqlite3.connect('CatLinh.sqlite')
cur = conn.cursor()

cur.execute('''
CREATE TABLE IF NOT EXISTS Orders(
    ID INTEGER PRIMARY KEY, Type_ID INTEGER, Container_ID TEXT,
    Customer_ID INTEGER, Destination_ID INTEGER, Vehicle INTEGER,
    PUL_ID INTEGER, DOL_ID INTEGER, From_Date INTEGER,
    To_Date INTEGER, Provision INTEGER, Incurred_Cost INTEGER, Description TEXT
);
''')

importdata()







