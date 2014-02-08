from math import sin, cos, pi
import simplekml

class RFKml():
    def __init__(self):
        """Using RF DataBase to create KML/KMZ files for Google Earth
        Some default parameters are set:
        Precious Degree = 5 (not suggest to change)
        Cell basic Radius = 1.265 km
        Radius Factor = 5 (Sector's radius / Circle's radius)
        Sector Beamwidth = 35        
        
        """
        self.preciousDeg = 5
        self.radius = 0.001265
        self.radiusFactor = 5
        self.beamwidth = 35
    
    def _calc_sec(self, _direction, _radius, _latlong):
        result = [0,0]
        try:
            result[0] = _latlong[0] + _radius * sin(_direction/180.0*pi)
            result[1] = _latlong[1] + _radius * cos(_direction/180.0*pi)
        except:
            print 'Calc Lat/long Error'
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
        
        
    def colorScheme(self):
        pass
    
        
        
        
        
        
        
        