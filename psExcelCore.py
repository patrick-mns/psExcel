import win32com.client
import os
import re
import json
import uuid
import pandas


class psExcelCore:
    def __init__(self, json, psd, dirsave):
        self.json = json
        self.psd = psd
        self.dirsave = dirsave
        
    def excelOpen(self):
        try:
            excel_data_fragment = pandas.read_excel(self.json)
            print(self.json)
            json_str = excel_data_fragment.to_json()
            data = json.loads(json_str)

            dataout = []
            x = 0

            while x < len(data[list(data.keys())[0]]):
                d = {}
    
                for key, value in data.items():
                    d.update({
                    key : value[str(x)]
                    })
    
                dataout.append(d)
                x += 1
            
            print(dataout)
            self.jsondata = dataout

        except:
            return(False)

    def existsPhotoshop(self):
        try:
            psApp = win32com.client.Dispatch("Photoshop.Application")
            self.psApp = psApp
            return(True)
        except:
            return(False)
    
    def psdOpen(self):
        try:
            self.psApp.Open(r"{}".format(self.psd))
            self.doc = self.psApp.Application.ActiveDocument
            return(True)
        except:
            return(False)

    def countLayers(self):
        i = 0
        while True:
            i += 1
            try:
                layer = self.doc.ArtLayers(i)
                print(layer)
            except:
                break
        self.lenlayers = i - 1
        return(True)
 
    def verifLayerText(self, id):
        try:
            layer = self.doc.ArtLayers(id)
            textLayer = layer.TextItem
            return(textLayer)
        except:
            return(False)
   
    def verifExistsKey(self, data_text):
        r = re.compile(r'\[[a-zA-Z\d]+\]')
        matches = r.findall(data_text.Contents)
        if matches:
           return(True)

    def dataLayers(self):
        dataLayers = []
        i = 0
        while self.lenlayers > i:
            data_text = self.verifLayerText(i)
            if data_text == False:
                None
            else:
                if self.verifExistsKey(data_text) == True:
                   dataLayers.append(data_text) 
            i += 1
        if len(dataLayers) > 0:
            self.layers = dataLayers
            return(True)
        else:
            return(False)

    def saveFilePNG(self):
        try:
            options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
            options.Format = 13
            options.PNG8 = False 
            pngfile = f"{self.dirsave}\\{uuid.uuid4()}.png"
            self.doc.Export(ExportIn=pngfile, ExportAs=2, Options=options)
            return(True)
        except:
            return(False)
        
    def start(self):
        if self.excelOpen()  == False: return('Invalid .xlsx file')
        if self.existsPhotoshop() == False: return('Photoshop not found')
        if self.psdOpen() == False: return('Invalid .psd file')
        if self.countLayers() != True: return('Invalid 0 layers')
        if self.dataLayers() != True: return('No text in the file\nor no [keys]')
        os.mkdir(self.dirsave) 
        for j in self.jsondata:
            baseLayer = []
            for layer in self.layers:
                l = layer.Contents
                baseLayer.append(l)
                for key in j.keys():
                    l = l.replace(f'[{key}]', str(j[key]))
                    layer.Contents = l
            if self.saveFilePNG() == False: return('Unable to save files')
            i = 0 
            for layer in self.layers:
                layer.Contents = baseLayer[i]
                i += 1
        return(True)