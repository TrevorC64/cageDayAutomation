from PIL import Image, ImageDraw
import pyqrcode as pyqr
import png
import openpyxl as xl
import pymysql

class CageDayGenerator:
    def __init__(self):
        self.logo = Image.open("logo.png")
        self.logo = self.logo.resize((75,30))
        self.sheet = ("inventory.xlsx")
        self.data = []
        self.connected = False


    ################################################
    ##              Excel Functions               ##
    ################################################   
    def grabData(self):
        self.wb = xl.load_workbook(self.sheet)
        self.ws = self.wb['Sheet1']
        for row in self.ws.iter_rows(min_row = 2, min_col=1, max_row=4,max_col=6):
            rrow = []
            for cell in row:
                rrow.append(cell.value)
            self.data.append(rrow)

    def showData(self):
        print(self.data)
        
    ################################################
    ##              Label Functions               ##
    ################################################  
    def createLabel(self, item):
        ## Input:
        ##   -> [name, id, desc, url, quantity]

        ## Generate QR Code
        self.qr = pyqr.create(item[3])
        self.qr.png('temp_qr.png', scale=2)
        self.qr = Image.open('temp_qr.png')

        ## Generate labels based off of total quantity:x
        for i in range(item[4]):               
            self.label = Image.new('RGB', (450,125), color = 'white')
            self.dlabel = ImageDraw.Draw(self.label)


            self.dlabel.text((10,10), 'Item:_________________________________________________', fill=(0,0,0))
            self.dlabel.text((50,10), item[0], fill=(0,0,0))

            self.dlabel.text((350,10), 'ID:_____________', fill=(0,0,0))
            self.dlabel.text((370,10), item[1], fill=(0,0,0))

            self.dlabel.text((10,30), 'Description:_________________________________________________', fill=(0,0,0))
            self.dlabel.text((90,30), item[2], fill=(0,0,0))
                
                
            self.dlabel.text((10,100), "Property of", fill=(0,0,0))
            self.label.paste(self.logo, (80,90))
            self.label.paste(self.qr, (360, 40))

                
            
            self.label.save(f'{item[0]}_{i}.png')

    
    def generateLabels(self):
        for i in self.data:
            self.createLabel(i)

    ################################################
    ##              SQL   Functions               ##
    ################################################
    def connect(self):                                                          #### TODO ####
        ## Connect to the SQL Database to be populated
        ## Sets the self.connected => true if sucessful
        host =""
        username="wiuxtcje_it"
        password=""
        db_name = "wiuxtcje_inventory"
        self.db = pymysql.connect(host, username, password, db_name)      ## Dont forget to call self.db.close() to close connection
        self.connected = True
        return 0

    def populate(self):                                                          #### TODO ####
        ## Populate the SQL Database
        self.connected = True
        if (self.connected):
            for i in self.data:
             ##  i -> [name, id, desc, url, quantity, price]
             ## INSERT INTO `wiux_inventory` (`name`, `id`, `description`, `price`, `count-in`, `count-total`) VALUES ('name', 'id', 'desc', '35', '2', '2');
                query = "INSERT INTO `wiux_inventory` (`name`, `id`, `description`, `price`, `count-in`, `count-total`) VALUES ('{name}', '{idd}', '{desc}', '{price}', '{c_in}', '{c_total}');".format(name=i[0],idd=i[1],desc=i[2],price=i[5],c_in=i[4],c_total=i[4])
                print(query)
                

            return 0
    
        else:
            print("Error! Not Connected to SQL DB, try self.connect()!")
                


        
x = CageDayGenerator()
x.grabData()
x.showData()
x.generateLabels()
#x.populate()
