__author__ = "fahadadeel"
import os.path
import sys

class Settings:
    """
        Add the jar to your path
    """
    sys.path.append("../../lib/aspose-cells-8.6.3.jar")
    
    
    dataDir = os.path.join(os.path.abspath("../../"), "data/")

    def __init__(self, dataDir):
        """
            : The path to the documents directory. :
        """
        
        Settings.dataDir = dataDir

        