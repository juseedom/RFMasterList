from xlrd import open_workbook
from xlutils.copy import copy
from difflib import get_close_matches
import logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

keywords = ('index','eNodeB ID_Sector No.','eNodeB Id','Lat','Long', 'Cell Id', 'Sector Id','EARFCN','TAL','TAI')


class RFDataBase():
    """Define a Object for RF Data Base
    
    Able to read RF data from Excel files;
    
    """
    def __init__(self, filepath):
        """
        For xlsx file, format is not read due to xlrd limit
        The default sheet index set to 0
        """
        try:
            if filepath.split(".")[-1] == "xls":
                rf_wb = open_workbook(filepath, formatting_info = True, on_demand = True)
                
            elif filepath.split(".")[-1] == "xlsx":
                rf_wb = open_workbook(filepath, formatting_info = False, on_demand = True)
            else:
                print "Unknown File Type."
                return False
            self.DB = dict()
            self.readFile(rf_wb)
                
        except TypeError:
            print "Cannot Find the Excel File."
        except IOError:
            print "Open Excel File Fialed."
        
        
            
            
    def readFile(self, rf_wb, sheet_index = 0):
        """Auto match the title rows according to keyword
        transfer RF data from workbook into dict"""
        rf_sheet = rf_wb.sheet_by_index(sheet_index)
        
        #find the title according to keywords from first 5 rows
        start_row = 0
        max_match = 0
        
        for i in range(5):
            s = 0
            str_match = [str(rf_sheet.row_values(i)[x]) for x in range(len(rf_sheet.row_values(i)))]
            for key in keywords:
                a = get_close_matches(key,str_match)
                s += len(a)
            if s > max_match:
                start_row = i
                max_match = s
        logging.debug("Found Title on row %i according to keyword match." %(start_row+1))
        
        #read title
        #match index or cell name and check duplicate
        str_title = rf_sheet.row_values(start_row)
        title_index = -1
        for i in ("index", "CellName"):
            a = get_close_matches(i, str_title)            
            if a:
                title_index = str_title.index(a[0])
                str_index = rf_sheet.col_values(title_index)[start_row+1:]                
                if [b for b in str_index if str_index.count(b)>1]:
                    title_index = -1
                    logging.debug("This Column %i cannot be index value as some are duplicated." %title_index)
                else:
                    logging.debug("Found Column %i can be index value" %title_index)
                    break
            else:
                logging.debug("Cannot find title name match with %s." %i)
                
        if title_index != -1:
            raw_data = dict()
            for i in range(start_row+1, rf_sheet.nrows):
                raw_value = rf_sheet.row_values(i)
                raw_data = dict(zip(str_title, raw_value))
                str_index = rf_sheet.cell(i, title_index).value
                self.DB[str_index] = raw_data
            logging.debug(len(self.DB))
        else:
            logging.debug("Cannot find the index")
        
         #format type (not suggested)
        
    def writeFile(self):
        pass
            
        
        


if __name__ == "__main__":
    test = RFDataBase("/home/juseedom/Downloads/RF_MasterList20140207.xls")
    

    
    