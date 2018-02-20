import jaydebeapi as sql # python to hsqldb connections

def readFlyStock(db_location,hsqldb_location):
    # Connect to db
    conn = sql.connect('org.hsqldb.jdbcDriver',
                    'jdbc:hsqldb:file:'+db_location,
                    ['sa',''],
                    hsqldb_location)
    curs = conn.cursor()
    curs.execute('SELECT * FROM \"Fly_Stock\"')
    result = curs.fetchall()
    results = '\"Stock_ID\",\"Genotype\",\"Description\",\"Note\",\"Res_Person\",\"Flybase\",\"Project\"'

    for tupe in result:
        results = results + '\n'
        for value in tupe:
            if value is not None:
                results = results + '\"' + value + '\"' + ','
            else:
                results = results + '\"' + 'None' + '\"' + ','

    results = results[:-1]
    results = results.encode('utf8')

    # Write to csv to prevent accessing database at every step
    file = open('Resources/LabDB.csv','wb')
    file.write(results)
    file.close()
    curs.close()
    conn.close()

def writeFlyStock(db_location, hsqldb_location, new_stocks):
    # Connect to db
    conn = sql.connect('org.hsqldb.jdbcDriver',
                    'jdbc:hsqldb:file:'+db_location,
                    ['sa',''],
                    hsqldb_location)
    curs = conn.cursor()

    # new_stocks is an array of tuples. HSQLDB 1.8 does not support multirow insert, thus the for loop.
    for x in new_stocks:
        curs.execute('INSERT INTO \"Fly_Stock\" (\"Stock_ID\", \"Genotype\", \"Description\", \"Note\", \"Res_Person\", \"Flybase\", \"Project\") VALUES ' + str(x))

    # Essentially rereads the database to update search results with new additions
    curs.execute('SELECT * FROM \"Fly_Stock\"')
    result = curs.fetchall()
    results = '\"Stock_ID\",\"Genotype\",\"Description\",\"Note\",\"Res_Person\",\"Flybase\",\"Project\"'
    for tupe in result:
        results = results + '\n'
        for value in tupe:
            if value is not None:
                results = results + '\"' + value + '\"' + ','
            else:
                results = results + '\"' + 'None' + '\"' + ','
    results = results[:-1]
    results = results.encode('utf8')

    # Replace LabDB.csv for proper search results
    file = open('Resources/LabDB.csv','wb')
    file.write(results)
    file.close()
    curs.close()
    conn.close()
