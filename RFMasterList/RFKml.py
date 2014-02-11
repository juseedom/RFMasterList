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
        self.radius = 0.001265
        self.radiusFactor = 5
        self.beamwidth = 35
        self.transparent = 6

        self.rules = dict()

    
    def _calc_sec(self, _direction, _radius, _latlong):
        result = [0,0]
        try:
            #Longitude, Latitude
            result[0] = _latlong[0] + _radius * sin(_direction/180.0*pi)
            result[1] = _latlong[1] + _radius * cos(_direction/180.0*pi)
        except:
            logging.debug('Calc Lat/long Error for %s %s %s' %(_direction, _radius, _latlong))
        finally:
            return result
        
    
    def cellShape(self, _shape, _direction, _latlong):
        """
        according to requested shape to return latlong boundary line
        """
        if _shape == 'Sector':
            tmp = []
            tmp.append(_latlong)
            for ang_step in range(0,self.beamwidth+self.preciousDeg,self.preciousDeg):
                tmp.append(self._calc_sec((ang_step+_direction-self.beamwidth/2)%360,self.radius,_latlong))
            tmp.append(_latlong)
            return tmp
        elif _shape == 'Circle':
            tmp = []
            for ang_step in range(0, 360, self.preciousDeg):
                tmp.append(self._calc_sec(ang_step,self.radius/self.radiusFactor,_latlong))
            return tmp
        else:
            return False


    def createCells(self, RFDB, map_title = None):
        kml = simplekml.Kml()

        #fol_site = kml.newfolder(name = 'Site')
        fol_cell = kml.newfolder(name = 'Cell')

        #fol_site.style = simplekml.Style(iconstyle = simplekml.IconStyle(scale = 0.4))
        #fol_site.style.iconstyle.icon.href = u'http://maps.google.com/mapfiles/kml/shapes/campground.png'

        for cell_index, cell_info in zip(RFDB.keys(),RFDB.values()):
            cell_type = dict()
            for rule_type, rule_value in zip(self.rules.keys(), self.rules.values()):
                if rule_type[0] == "PCI":
                    cell_type[rule_type[1]] = rule_value[str(int(cell_info[rule_type[0]])%3)]
                elif rule_type[0] == "EARFCN":
                    cell_type[rule_type[1]] = rule_value[str(cell_info[rule_type[0]]).split(".")[0]]
                else:
                    cell_type[rule_type[1]] = rule_value[cell_info[rule_type[0]]]
            #logging.debug("Create cell type for %s : %s" %(cell_index, cell_type))


            if cell_info["Azimuth"] and cell_info["Latitude"] and cell_info["Longitude"]:
                draw_cell = fol_cell.newpolygon(name = cell_index, \
                    outerboundaryis = self.calc_latlong(cell_type["Shape"], cell_info["Azimuth"], (cell_info["Longitude"],cell_info["Latitude"])),\
                    description = "\n".join(str(cell_info).split(",")) \
                    )
                transferColor = lambda s: "ff"+s[1:] if s.find("#") != -1 else getattr(simplekml.Color, s)
                draw_cell.style = simplekml.Style(linestyle = simplekml.LineStyle(width =2, color = str(self.transparent)+transferColor(cell_type["Line Color"])[1:]),\
                                polystyle = simplekml.PolyStyle(color =str(self.transparent)+ transferColor(cell_type["Fill Color"])[1:]))
                
            else:
                logging.debug("Missing information for cell %s, create cell failed." % cell_index)

        kml.savekmz('D:\\123.kmz')

    def calc_latlong(self, _shape = 'Circle', _direction = 0, _latlong = None):
        if _shape == 'Sector':
            tmp = []
            tmp.append(_latlong)
            for ang_step in range(0,self.beamwidth+self.preciousDeg,self.preciousDeg):
                tmp.append(self._calc_sec((ang_step+_direction-self.beamwidth/2)%360,self.radius,_latlong))
            tmp.append(_latlong)
            return tmp
        elif _shape == 'Circle':
            tmp = []
            for ang_step in range(0, 360, self.preciousDeg):
                tmp.append(self._calc_sec(ang_step,self.radius/5,_latlong))
            return tmp
        else:
            return False

    def createRules(self, addrules):
        for rule_type, rule_value in zip(addrules.keys(), addrules.values()):
            rf_type = rule_type[0]
            kml_type = rule_type[1]
            rf_value = rule_value.keys()
            kml_value = rule_value.values()
            print rf_type, kml_type, rf_value, kml_value
            if kml_type in ("Fill Color", "Line Color", "Shape"):
                self.rules.update({(rf_type, kml_type):dict(zip(rf_value, kml_value))})
                logging.debug("Add rules successfully for %s and %s." %(rf_type, kml_type))
            else:
                logging.debug("This option is not support now.")
                return False


if __name__ == '__main__':
    rf_kml = RFKml()
    test = RFDataBase()
    test.readFile(".\\RF_MasterList20140207.xls")
    test.readSheet(0)
    rules = {("EARFCN", "Fill Color"):{"1850":"red", "2970":"blue", "1800":"green"},
            ("Type", "Shape"):{"Indoor":"Circle", "Outdoor":"Sector"},
            ("PCI", "Line Color"):{"1":"blue", "2":"#dc143c", "0":"yellow"}
            }
    rf_kml.createRules(rules)

    rf_kml.createCells(test.DB)
