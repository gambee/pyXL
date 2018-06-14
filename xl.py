import xlsxwriter

def xlTable(sheet, data, coord=(0, 0), name=None):
    # sheet is the xlsxwriter.Worksheet object 
    # data can at least at this point, be a dict
    # with a columns member and a rows member,
    # or a columns member and a records member
    # Returns the coordinates of the bottom right corner.

    # Proess and validate arguments
    if('columns' in data):
        cols = data['columns']
        if('rows' in data):
            rows = data['rows']
        elif('records' in data):
            rows = [
                [
                    d[k] if k in d else None
                    for k in cols
                ]
                for d in data['records']
            ]
        else:
            raise Exception("xlTable expects 'rows' key or 'records' key")
    else:
        raise Exception("xlTable expects 'columns' key")

    options = {
        'data': rows,
        'columns':[{'header': col} for col in cols]
        }
    if(name is not None):
        options['name'] = name
    
    ecoord = (coord[0] + len(rows), coord[1] + len(cols) - 1)

    # add the table on the data just written
    sheet.add_table(coord[0],
                    coord[1],
                    ecoord[0],
                    ecoord[1],
                    options)

    return ecoord
