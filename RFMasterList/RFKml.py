from math import sin, cos, pi
import simplekml

import sys, os
sys.path.append(os.path.abspath(os.curdir))
from RFDB import RFDataBase


import logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

class RFKml():
    def __init__(self):
        """Using RF DataBase to create KML/KMZ files for Google Earth
        Some default parameters are set:
        Precious Degree = 5 (not suggest to change)
        Cell basic Radius = 1.265 km
        Radius Factor = 5 (Sector's radius / Circle's radius)
        Sector Beamwidth = 35  
        Cell color transparent = 6 (range from 0~F)      
        
        """
        self.preciousDeg = 5
        #self.radius = 0.001265
        #self.radiusFactor = 5
        self.beamwidth = 35
        self.transparent = 6

        self.rules = dict()
        
        self.supportOptions = (("EARFCN", "PCI", "Type"), ("Shape", "Fill Color", "Line Color"))
        #self.supportOptions = {"Shape":("Circle","Sector)")}

    
    def _calc_sec(self, _direction, _radius, _latlong):
        try:
            #Longitude, Latitude
            a = _latlong[0] + _radius * sin(_direction/180.0*pi)
            b = _latlong[1] + _radius * cos(_direction/180.0*pi)
        except:
            logging.debug('Calc Lat/long Error for %s %s %s' %(_direction, _radius, _latlong))
        finally:
            return (a,b)
        
    def createCells(self, RFDB, map_title = None):
        kml = simplekml.Kml()
        #fol_site = kml.newfolder(name = 'Site')
        fol_cell = kml.newfolder(name = 'Cell')
        #fol_site.style = simplekml.Style(iconstyle = simplekml.IconStyle(scale = 0.4))
        #fol_site.style.iconstyle.icon.href = u'http://maps.google.com/mapfiles/kml/shapes/campground.png'

        for cell_index, cell_info in zip(RFDB.keys(),RFDB.values()):
            #create cell types cell_type["Shape"] = "Sector", "0.1", "10"
            cell_type = dict()
            for rule_type, rule_value in zip(self.rules.keys(), self.rules.values()):
                if rule_type[0] == "PCI":
                    cell_type[rule_type[1]] = rule_value[str(int(cell_info[rule_type[0]])%3)]
                elif rule_type[0] == "EARFCN" or rule_type[0] == "Azimuth":
                    cell_type[rule_type[1]] = rule_value[str(cell_info[rule_type[0]]).split(".")[0]]
                else:
                    cell_type[rule_type[1]] = rule_value[cell_info[rule_type[0]]]
            #logging.debug("Create cell type for %s : %s" %(cell_index, cell_type))

            #need to add mat_table here
            if cell_info["Latitude"] and cell_info["Longitude"]:
                if cell_type["Height"] == "0":
                    draw_cell = fol_cell.newpolygon(name = cell_index,\
                        outerboundaryis = self.calc_latlong(cell_type["Shape"],cell_type["Cell Radius"],"0",int(cell_info["Azimuth"]), (cell_info["Longitude"],cell_info["Latitude"])),\
                        description = "\n".join(str(cell_info).split(",")))                   
                else:
                    draw_cell = fol_cell.newpolygon(name = cell_index, altitudemode = 'relativeToGround', extrude = 1,\
                        outerboundaryis = self.calc_latlong(cell_type["Shape"],cell_type["Cell Radius"], cell_type["Height"], int(cell_info["Azimuth"]), (cell_info["Longitude"],cell_info["Latitude"])),\
                        description = "\n".join(str(cell_info).split(",")))
                transferColor = lambda s: simplekml.Color.rgb(int(s[1:3],16), int(s[3:5],16), int(s[5:7],16)) if s.find("#") != -1 else getattr(simplekml.Color, s)
                draw_cell.style = simplekml.Style(linestyle = simplekml.LineStyle(width =2, color = str(self.transparent)+transferColor(cell_type["Line Color"])[1:]),\
                                polystyle = simplekml.PolyStyle(color = str(self.transparent) + transferColor(cell_type["Fill Color"])[1:]))
               
            else:
                logging.debug("Missing information for cell %s, create cell failed." % cell_index)

        kml.savekmz('123.kmz')
        logging.debug("Create KML file successfully")

    def calc_latlong(self, _shape ='Circle',_cellRadius = '0.001265',_height = "0", _direction = 0, _latlong = None):
        #print _latlong
        if _shape == 'Sector':
            tmp = list()
            tmp.append(_latlong)
            for ang_step in range(0,self.beamwidth+self.preciousDeg,self.preciousDeg):
                tmp.append(self._calc_sec((ang_step+_direction-self.beamwidth/2)%360,float(_cellRadius),_latlong))
            tmp.append(_latlong)
        elif _shape == 'Circle':
            tmp = []
            for ang_step in range(0, 360, self.preciousDeg):
                tmp.append(self._calc_sec(ang_step,float(_cellRadius),_latlong))
        else:
            return False
        if _height != "0":
            tmp = [a+(_height,) for a in tmp]
        return tmp

    def createRules(self, addrules):
        for rule_type, rule_value in zip(addrules.keys(), addrules.values()):
            rf_type = rule_type[0]
            kml_type = rule_type[1]
            rf_value = rule_value.keys()
            kml_value = rule_value.values()
            print rf_type, kml_type, rf_value, kml_value
            if kml_type in ("Fill Color", "Line Color", "Shape", "Cell Radius", "Height"):
                self.rules.update({(rf_type, kml_type):dict(zip(rf_value, kml_value))})
                logging.debug("Add rules successfully for %s and %s." %(rf_type, kml_type))
            else:
                logging.debug("This option is not support now.")
                return False


if __name__ == '__main__':
    rf_kml = RFKml()
    test = RFDataBase()
    test.readFile(".\\Test File\\RF_MasterList20140207.xls")
    test.readTitle()
    test.readSheet({"Index":"Index", "EARFCN":"EARFCN", "Type":"Type", "PCI":"PCI"})
    rules = {("EARFCN", "Fill Color"):{"1850":"red", "2970":"blue", "1800":"green"},
            ("Type", "Shape"):{"Indoor":"Circle", "Outdoor":"Sector"},
            ("PCI", "Line Color"):{"1":"blue", "2":"#dc143c", "0":"yellow"},
            ("EARFCN", "Height"):{"1850":"0", "2970":"0", "1800":"0"}
            }
    rf_kml.createRules(rules)

    rf_kml.createCells(test.DB)
