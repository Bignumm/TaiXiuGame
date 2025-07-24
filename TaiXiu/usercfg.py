from openpyxl import load_workbook

def HAdmin():
    options = input('p/c/w/m/per/q | (Play/CreateAcc/Wr/Money/Permission/quit): ')
    while ( options != 'p' and options != 'c' and options != 'w' and options != 'm' and options != 'per' and options != 'q'): 
        options = input('p/c/w/m/per/q (Play/CreateAcc/Wr/Money/Permission/quit) | Please input again: ')

    if(options == 'q'):
        return False
    
    if(options == 'p'):
        import gp
        gp.game(2)
        return

    if(options == 'c'): 
        wb = load_workbook('Data.xlsx')
        ws = wb.active
        count = ws['E2'].value + 1
        ws['A' + str(count)] = count - 1
        ws['B' + str(count)] = input('Input username: ')
        ws['C' + str(count)] = input('Input password: ')
        bal = input('Bal: ')
        while not bal.isdigit():
            bal = input('Input again: ')
        ws['D' + str(count)] = bal
        per = int(input('1/2 | Admin/User: '))
        while 1 < per > 2: 
            options = input('input C/Q (continue/quit): ')
            if(options == 'q' or options == 'Q'):
                return
            per = input('Input again: ')
        ws['F' + str(count)] = per
        print('The Account has been created succesfully')
        ws['E2'] = count
        wb.save('Data.xlsx')
        return
    
    if options == 'w': 
        wr = int(input('(0/1/6) | Config wr | Normal/Xiu/Tai: '))
        while wr != 0 and wr != 1 and wr != 6: 
            options = input('input C/Q (continue/quit): ')
            if(options == 'q' or options == 'Q'):
                return
            wr = input('(0/1/6) | Config wr: ')
        print("Config successfully")
        wb = load_workbook('Data.xlsx')
        ws = wb.active
        ws['H1'] = wr
        wb.save('Data.xlsx')
        return
    
    if options == 'm':
        wb = load_workbook('Data.xlsx')
        ws = wb.active
        count = ws['E2'].value + 1
        username = input('Nhap ten: ')
        for i in range (2, count + 1, 1):
            if(username == ws['B' + str(i)].value):
                bal = int(input('Bal: '))
                ws['D' + str(i)] = (bal + int(ws['D' + str(i)].value))
                if bal > 0: print('Add successfully!')
                else: print('Sub successfully!')
                wb.save('Data.xlsx')
                return
        
    
    if options == 'per':
        wb = load_workbook('Data.xlsx')
        ws = wb.active
        count = ws['E2'].value + 1 
        count = ws['E2'].value + 1
        for i in range (2, count, 1):
            print('%s | %s'%(ws['B' + str(i)].value, ws['F' + str(i)].value ))
        wb.save('Data.xlsx')  
        
        username = input('Username: ')
        per = int(input('Per: '))
        while per == 0 and per != 1 and per != 2:
            options = input('input C/Q (continue/quit): ')
            if(options == 'q' or options == 'Q'):
                return  
            per = int(input('Input again: '))
        for i in range(2, count, 1):
            wb = load_workbook('Data.xlsx')
            ws = wb.active
            if username == ws['B' + str(i)].value:
                ws['F' + str(i)] = per
                print('Change successfully')
                wb.save('Data.xlsx')
                return


def Admin(id):
    options = input('p/c/w/m/q | (Play/CreateAcc/Wr/Money/quit): ')
    while ( options != 'p' and options != 'c' and options != 'w' and options != 'm' and options != 'per'): 
        options = input('p/c/w/m/q (Play/CreateAcc/Wr/Money/quit | Please input again: ')

    if(options == 'q'):
        return False
    
    if(options == 'p'):
        import gp
        gp.game(id)
        return

    if(options == 'c'): 
        wb = load_workbook('Data.xlsx')
        ws = wb.active
        count = ws['E2'].value + 1
        ws['A' + str(count)] = count - 1
        ws['B' + str(count)] = input('Input username: ')
        ws['C' + str(count)] = input('Input password: ')
        bal = input('Bal: ')
        while not bal.isdigit():
            options = input('input C/Q (continue/quit): ')
            if(options == 'q' or options == 'Q'):
                return
            bal = input('Input again: ')
        ws['D' + str(count)] = bal
        per = int(input('1/2 | Admin/User: '))
        while 1 < per > 2: 
            options = input('input C/Q (continue/quit): ')
            if(options == 'q' or options == 'Q'):
                return
            per = input('Input again: ')
        ws['F' + str(count)] = per
        print('The Account has been created succesfully')
        ws['E2'] = count
        wb.save('Data.xlsx')
        return
    
    if options == 'w': 
        wr = int(input('(0/1/6) | Config wr | Normal/Xiu/Tai: '))
        while wr != 0 and wr != 1 and wr != 6: 
            options = input('input C/Q (continue/quit): ')
            if(options == 'q' or options == 'Q'):
                return
            wr = input('(0/1/6) | Config wr: ')
        print("Config successfully")
        wb = load_workbook('Data.xlsx')
        ws = wb.active
        ws['H1'] = wr
        wb.save('Data.xlsx')
        return
    
    if options == 'm':
        wb = load_workbook('Data.xlsx')
        ws = wb.active
        count = ws['E2'].value + 1
        username = input('Nhap ten: ')
        for i in range (2, count + 1, 1):
            if(username == ws['B' + str(i)].value):
                bal = int(input('Bal: '))
                ws['D' + str(i)] = (bal + int(ws['D' + str(i)].value))
                if bal > 0: print('Add successfully!')
                else: print('Sub successfully!')
                wb.save('Data.xlsx')
                return

def user(id):
    options = input('p/m/q | (Play/Money/quit): ')
    while ( options != 'p' and options != 'c' and options != 'w' and options != 'm' and options != 'per'): 
        options = input('p/m (Play/Money) | Please input again: ')

    if(options == 'q'):
        return False
    
    if(options == 'p'):
        import gp
        gp.game(id)
        return
    
    if options == 'm':
        wb = load_workbook('Data.xlsx')
        ws = wb.active
        bal = int(input('Bal: '))
        if ws['D' + str(id)].value < 0 and bal < 0: 
            print('Sub fail')
            return
        ws['D' + str(id)] = (bal + int(ws['D' + str(id)].value))
        if bal > 0: print('Add successfully!')
        else: print('Sub successfully!')
        wb.save('Data.xlsx')
        return