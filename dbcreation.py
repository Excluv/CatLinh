import sqlite3
import pandas as pd

conn = sqlite3.connect('CatLinh.sqlite')
cur = conn.cursor()

# Create all necessary tables: unfinished
cur.executescript('''
CREATE TABLE IF NOT EXISTS Type(
    ID INTEGER PRIMARY KEY, Name TEXT UNIQUE
);
CREATE TABLE IF NOT EXISTS Customer(
    ID INTEGER PRIMARY KEY, Name TEXT UNIQUE
);
CREATE TABLE IF NOT EXISTS PickupLocation(
    ID INTEGER PRIMARY KEY, Name TEXT UNIQUE, Type_ID INTEGER, Cost INTEGER
);
CREATE TABLE IF NOT EXISTS DropoffLocation(
    ID INTEGER PRIMARY KEY, Name TEXT UNIQUE, Type_ID INTEGER, Cost INTEGER
);
CREATE TABLE IF NOT EXISTS Destination(
    ID INTEGER PRIMARY KEY, Name TEXT UNIQUE
);
CREATE TABLE IF NOT EXISTS Revenue(
    Type_ID INTEGER, Customer_ID INTEGER, Destination_ID INTEGER, Revenue INTEGER,
    UNIQUE(Type_ID, Customer_ID, Destination_ID)
);
CREATE TABLE IF NOT EXISTS Salary(
    Vehicle INTEGER, Destination_ID INTEGER, Type_ID INTEGER, Salary INTEGER,
    UNIQUE(Vehicle, Destination_ID, Type_ID, Salary)
);
''')


def file_to_db(name):
    # Open excel file to input information into the database:
    df = pd.read_excel(f'C:\\Users\\jio\\Desktop\\CatlinhProject\\Database Constructor\\{name}.xlsx')
    data = list(map(tuple, df.values))

    # Retrieve the column names from the database:
    cur.execute(f'SELECT * FROM {name}')
    names = list(map(lambda title: title[0], cur.description))

    # Structure the command line:
    if names[0] == 'ID':
        # Ignore ID column (auto incrementing primary key specified):
        values = f'(?{", ?" * (len(names)-2)})'
    else:
        # In case a table does not have 'ID' column, every column needs to be filled
        values = f'(?{", ?" * (len(names)-1)})'

    command = f'INSERT OR IGNORE INTO {name} ('
    for index in range(len(names)):
        if names[index] == 'ID': continue
        if index == (len(names) - 1):
            command = command + f'{names[index]}'
        else:
            command = command + f'{names[index] + ", "}'
    command = command + f') VALUES {values}'

    # Insert actual data into the database:
    cur.executemany(command, data)
    conn.commit()


file_to_db('Customer')
file_to_db('Destination')
file_to_db('Type')
file_to_db('Revenue')
file_to_db('PickupLocation')
file_to_db('DropoffLocation')
file_to_db('Salary')
cur.close()