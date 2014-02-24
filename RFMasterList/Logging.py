import ConfigParser
import os

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
            self.config = ConfigParser.RawConfigParser()
            self.config.read("Excel2KML.cfg")
            self.config.get("section", "option")
            
            
        else:
            self.config = ConfigParser.RawConfigParser()
            self.config.add_section("section")
            self.config.set("section", "option", "value")
                
                
    def remember_cfg(self):
        with open("Excel2KML.cfg", "wb") as configfile:
            self.config.write(configfile)
    
    
    
