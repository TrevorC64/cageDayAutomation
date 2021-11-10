from PIL import Image, ImageDraw, ImageFont
import pyqrcode as pyqr
import png
import openpyxl as xl
import pymysql

class CageDayGenerator:
    def __init__(self):
        self.logo = Image.open("logo.png")
        self.logo = self.logo.resize((75,30))
        self.sheet = "inventory.xlsx"
        self.data = []


    ################################################
    ##              Excel Functions               ##
    ################################################   
    def grabData(self):
        self.wb = xl.load_workbook(self.sheet)
        self.ws = self.wb['Sheet1']
        for row in self.ws.iter_rows(min_row = 2, min_col=1, max_row=115,max_col=12):
            if row[9].value: # True or False flag if the item will be tagged
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
        ##   -> [name, id_num, dept, location, desc, url, quantity, price, labelType, willTag?, brand, serial]


        ## There are 4 different label types:
        ##      o  1 -> bv ~ Big Vertical
        ##      o  2 -> sv ~ Small Vertical
        ##      o  3 -> bh ~ Big Horizontal
        ##      o  4 -> sh ~ Small Horizontal

        font = ImageFont.truetype("bsbold.otf", 64)
                                
        ## Generate QR Code
        iid = str(item[1])
        dept = item[2]
        if dept == 'audio':
            iid = 'aud' + iid
        elif dept == 'visual':
            iid = 'vis' + iid
        elif dept == 'eng':
            iid = 'eng' + iid
        elif dept == 'ops':
            iid = 'ops' + iid
        elif dept == 'it':
            iid = 'it' + iid
                
        qr = pyqr.create("https://wiux.org/inventory/?i_id=" + iid)
        qr.png('temp_qr.png', scale=8)
        qr = Image.open('temp_qr.png')
 
        for i in range(item[6]):
            print(i)
            label_type = 1
            name = item[0] + " (" + item[3] + ")"
            print(name)
            
            desc = item[4]
            if desc == None:
                desc = ""
            
            loc = item[3]
            
            
            if label_type == 1:
                with Image.open('labels/tag_bv.png') as label:
                    draw = ImageDraw.Draw(label)
                    draw.text((350,670), name, fill=(0,0,0), font=font)
                    draw.text((300, 1690), iid, fill=(0,0,0), font=font)


                    # Max description length : 58
                    lines = self.formatDesc(desc, (20, 38))
                    
                    draw.text((700,1140), lines[0], fill=(0,0,0), font=font)
                    draw.text((75,1400), lines[1], fill=(0,0,0), font=font)

                    label.paste(qr, (1100,25))
                    
            elif label_type == 2:
                with Image.open('labels/tag_sv.png') as label:
                    draw = ImageDraw.Draw(label)
                    draw.text((290,575), name, fill=(0,0,0), font=font)
                    draw.text((250,1525), iid, fill=(0,0,0), font=font)

                   # Max description lenght : 44

                    lines = self.formatDesc(desc, (19,25))
                    draw.text((500,975), lines[0], fill=(0,0,0), font=font)
                    draw.text((110, 1250),lines[1], fill=(0,0,0), font=font)
                    
                    label.paste(qr, (675, 1725))
                
            elif label_type == 3:
                with Image.open('labels/tag_bh.png') as label:
                    draw = ImageDraw.Draw(label)
                    draw.text((250,450), name, fill=(0,0,0), font=font)
                    draw.text((225, 815), iid, fill=(0,0,0), font=font)
                    
                   # Max description lenght : 44

                    draw.text((475,640), desc, fill=(0,0,0), font=font)

                    label.paste(qr, (1700, 25))
                    
            elif label_type == 4:
                with Image.open('labels/tag_sh.png') as label:
                    draw = ImageDraw.Draw(label)
                    draw.text((350,480), name, fill=(0,0,0), font=font)
                    draw.text((250, 1100), iid, fill=(0,0,0), font=font)

                    #Max description length : 37
    
                    draw.text((725,790), desc, fill=(0,0,0), font=font)

                    label.paste(qr, (1700,25))
                    
            else:
                print(f'Invalid Label Format for {name}, given {label_type}, expected (1,2,3,4)')
            label.save(f'out/{item[0]}_{i}.png')
    
                   
    def formatDesc(self, string, lengths):
        # returns an array of size lengths with substrings where each string is less than its corresponding lengths
        ss = string.split()
        out = []
        
        # For each line
        for length in lengths:
            currentString = ""
            currentLength = 0
            while currentLength < length and len(ss) > 0:
                if currentLength + len(ss[0]) + 1 < length:
                    currentString += ss[0] + " "
                    if len(ss) == 1:
                        ss=[]
                        break
                    ss = ss[1:]
                    currentLength += len(ss[0]) + 1
                else:
                    break
              
            out += [currentString]

        return out
        

            
    
    def generateLabels(self):
        for i in self.data:
            self.createLabel(i)            

        
x = CageDayGenerator()

x.grabData()
x.generateLabels()

## If we want to automate adding labels to a word doc and resizing we can create a
## dictionary name->size mapping file names
## (truncating the number) to the respective sizes (width, height) in px/in??
## that we want the labels to be
##   Algo:
##     1. Import all image file names
##     2. Iterate over that list of imageNames
##          i. Lookup imageName in dict name->size
##         ii. Paste the image into word doc using add_image?
##               - Maybe add boarders to each to make cutting easier, and/or make more asthetically pleasing
##        iii. Possibly add logic to minimize the number of pages needed to print
##             ### NOTE: there is a lot of areas to optimize, such as sorting the list of images to fit the maximum
##                       area of 'label' per page



