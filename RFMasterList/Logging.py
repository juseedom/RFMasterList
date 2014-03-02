import ConfigParser
import os
import json

import logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

class RFLogging():
    def __init__(self):
        """
        Check cfg file existed or not:
        Existed - read CFG;
        NOT existed - create a CFG with initial values
        """
        #currentpath = os.path.dirname(os.path.abspath(__file__))
        if os.path.isfile("Excel2KML.cfg"):
            #read configuration
            with open("Excel2KML.cfg", "rb") as configfile:
                self.readRuleFromCfg(configfile)

        else:
            self.createDefaultRule()
            self.writeRuleToCfg()
         
         
    def createDefaultRule(self):
        #build default rules
        rf_item = dict()
        #read RFDB
        
        #change to tab_setting, do KPI statistics now
        rf_item["EARFCN"] = ["1850", "2970"]#self.RFDB.DBStatistics("EARFCN")
        rf_item["Type"] = ["Indoor", "Outdoor"]#self.RFDB.DBStatistics("Type")
        rf_item["PCI"] = ['0', '1', '2']

        kml_item = dict()
        kml_item["Shape"] = ["Circle", "Sector"]
        kml_item["Cell Radius"] = ["0.00025", "0.001265"]
        kml_item["Height"] = list()
        #kml_item["Shape"] = [["Circle","0.00025","20"],["Sector","0.001265","10"]]
        kml_item["Line Color"] = list()
        kml_item["Fill Color"] = list()

        self.rules = dict()
            
        self.rules[("EARFCN", "Fill Color")] = dict().fromkeys(rf_item["EARFCN"], "blue")
        self.rules[("PCI", "Line Color")] = dict().fromkeys(rf_item["PCI"], "red")
        self.rules[("Type", "Shape")] = dict(zip(rf_item["Type"],kml_item["Shape"]))
        self.rules[("Type", "Cell Radius")] = dict(zip(rf_item["Type"],kml_item["Cell Radius"]))
        self.rules[("EARFCN", "Height")] = dict().fromkeys(rf_item["EARFCN"], "0")
    
    def writeRuleToCfg(self):
        try:
            self.config = ConfigParser.RawConfigParser()
            self.config.add_section("RuleType")
            self.config.add_section("RuleValue")
            for i, type_list in enumerate(self.rules):
                self.config.set("RuleType", "type%i" %i, json.dumps(type_list))
                self.config.set("RuleValue", "value%i" %i,json.dumps(self.rules.values()[i]))
            with open("Excel2KML.cfg", "wb") as configfile:
                self.config.write(configfile)
            logging.debug("Saving the CFG successfully")
            return True
        except:
            logging.debug("Write Rule to CFG file failed, pls check CFG file")
            return False            
                
        
                
    def readRuleFromCfg(self, configfile):
        try:
            self.config = ConfigParser.RawConfigParser()
            self.config.readfp(configfile)        
            self.rules = dict()
            for i in range(len(self.config.options("RuleType"))):
                rule_type = tuple(json.loads(self.config.get("RuleType", "type%i" %i)))
                self.rules[rule_type] = json.loads(self.config.get("RuleValue", "value%i" %i))
            logging.debug("Read rule from CFG successfully")
            return self.rules
        except:
            logging.debug("Read Rule From CFG file failed, pls check CFG file or delete it")
            return False
            
if __name__ == "__main__":
    test = RFLogging()
    
    
    
    
    
    
    
