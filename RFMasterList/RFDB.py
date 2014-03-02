from xlrd import open_workbook
from xlutils.copy import copy
from difflib import get_close_matches
import logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


class RFDataBase():
    """Define a Object for RF Data Base
    
    Able to read RF data from Excel files;
    
    """
    def __init__(self):
        self.DB = dict()
        self.RFIndex = ['Index','eNodeB Id','Latitude','Longitude','Sector Id','Cell Id','EARFCN','Azimuth','Type','PCI', 'TAI']
        self.tableTitle = dict()

            
    def readFile(self, filepath):
        """
        For xlsx file, format is not read due to xlrd limit
        The default sheet index set to 0
        """
        try:
            if filepath.split(".")[-1] == "xls":
                self.rf_wb = open_workbook(filepath, formatting_info = True, on_demand = True)
            elif filepath.split(".")[-1] == "xlsx":
                self.rf_wb = open_workbook(filepath, formatting_info = False, on_demand = True)
            else:
                logging.debug("Unknown File Type.")
                self.rf_wb = False              
        except TypeError:
            logging.debug("Cannot Find the Excel File.")
        except IOError:
            logging.debug("Open Excel File Fialed.")
        finally:
            if self.rf_wb:
                return self.rf_wb.sheet_names()
            else:
                return False

    def readTitle(self, sheet_index = 0):
        """Auto match the title rows according to keyword
        transfer RF data from workbook into dict"""
        self.rf_sheet = self.rf_wb.sheet_by_index(sheet_index)
        
        #find the title according to keywords from first 5 rows
        self.start_row = 0
        max_match = 0
        
        for i in range(5):
            s = 0
            str_match = [str(self.rf_sheet.row_values(i)[x]) for x in range(len(self.rf_sheet.row_values(i)))]
            for key in self.RFIndex:
                a = get_close_matches(key,str_match)
                s += len(a)
            if s > max_match:
                self.start_row = i
                max_match = s

        logging.debug("Found Title on row %i according to keyword match." %(self.start_row+1))
        self.str_title = self.rf_sheet.row_values(self.start_row)
        
        return self.str_title
        
    def readSheet(self, table_title):
        #read title
        #match index or cell name and check duplicate
        #transfer all number to float, str otherwise 
        self.tableTitle = table_title      
        title_index = self.str_title.index(table_title["Index"])
        str_index = self.rf_sheet.col_values(title_index)[self.start_row+1:]
        if [b for b in str_index if str_index.count(b)>1]:
            title_index = -1
            logging.debug("This Column %i cannot be index value as some are duplicated." %table_title["Index"])
            return False
        else:
            logging.debug("Selected Column %s can be index value" %table_title["Index"])   
        
        logging.debug("Start to read RF Data into DB")     
        raw_data = dict()
        for i in range(self.start_row+1, self.rf_sheet.nrows):
            raw_value = self.rf_sheet.row_values(i)
            raw_value = [float(a) if str(a).replace(".","",1).isdigit() else str(a) for a in raw_value]
            raw_data = dict(zip(self.str_title, raw_value))
            str_index = self.rf_sheet.cell(i, title_index).value
            self.DB[str_index] = raw_data

        logging.debug("Load %i LTE Cells" %len(self.DB))

        #format type (not suggested)
    
       


    def DBStatistics(self, dedicate_index):
        """
        Remove all those duplicated value for specific Parameters
        1st Step: remove duplicated objects
        2nd Step: remove duplicated values

        some will transfer value type while all others remain str:
        EARFCN->str->float

        """

        logging.debug("Start to statistics for %s" %dedicate_index)
        table_index = self.tableTitle[dedicate_index]
        dedicate_value = set([raw_data[table_index] for raw_data in self.DB.values()])
        

        if dedicate_index in ("EARFCN"):
            logging.debug("Start to transfer All EARFCN value into int")
            #dedicate_value = set([float(i) for i in dedicate_value])
            dedicate_value = set([str(int(i)) for i in dedicate_value])
            
        return dedicate_value

    def Update2File(self):
        pass
            
        
        


if __name__ == "__main__":
    test = RFDataBase()
    test.readFile('/home/juseedom/Downloads/RF_MasterList20140207.xls')
    test.readTitle(0)
    test.readSheet({"Index":"index","EARFCN":"EARFCN"})
    #test = RFDataBase("D:\\CODES\\Python\\RFMasterList\\RFMasterList\\RFMasterList\\Test File\\RF_MasterList20140207.xls")
    #print test.DBStatistics("SiteAddress")
    

    
    
