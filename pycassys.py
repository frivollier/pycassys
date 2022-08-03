'''
@author: frivollier

pycasssy.py - module to interact with CASSYS

2018.08.30 Initial draft

Pre-requisites:
    CASSYS_Engine.exe should be in lib folder TODO: add to
    xmltodict #http://docs.python-guide.org/en/latest/scenarios/xml/ #https://github.com/martinblech/xmltodict
    pandas

Overview:

'''
'''
Revision history

0.0.1:  Initial stable release
'''

'''
%load_ext autoreload
%autoreload 2
'''

import os, sys, datetime, shutil
import xmltodict #http://docs.python-guide.org/en/latest/scenarios/xml/ #https://github.com/martinblech/xmltodict
import subprocess
import pandas as pd
import tempfile

import pkg_resources
global CASSYS_PATH # path to CASSYS EXE
CASSYS_PATH = os.path.abspath(pkg_resources.resource_filename('pycassys', 'bin/') )

import clr
CASSY_DLL = os.path.join(CASSYS_PATH, 'CASSYS.dll')
clr.AddReference(CASSY_DLL)
from CASSYS import CASSYSClass

class cassysObj:
    '''
    cassysOBj:  top level class to work with CASSSY,

    values:

        path        : working directory with Radiance materials and objects
        TODO:  populate this more
    functions:
        __init__   : initialize the object
        _setPath    : change the working directory
        TODO:  populate this more

    '''

    def __init__(self, csyx_file, climate_file, output_folder=None):
        '''
        Description
        -----------
        initialize cassysOBJ

        Parameters
        ----------
        csyx_file: string, path to .csyx file
        climate_file:
        output_folder:

        Returns
        -------
        none
        '''
        #class variables
        self.dict = dict()        # create empty dict to load csyx XML


        now = datetime.datetime.now()
        self.nowstr = str(now.date())+'_'+str(now.hour)+str(now.minute)+str(now.second)

        ''' DEFAULTS '''

        #Constant
        self.cassys_exe = os.path.join(CASSYS_PATH,'CASSYS Engine.exe')
        self.errorlog = os.path.join(CASSYS_PATH,'ErrorLog.txt')

        #Variables
        self.config_file = csyx_file
        self.climate_file = climate_file

        if output_folder is None:
            self.output_path = tempfile.gettempdir()
        else:
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            self.output_path= output_folder

        ''' Read csyx'''
        #open cassys csyx config file into a dictionary
        with open(self.config_file) as fd:
            self.csyx = xmltodict.parse(fd.read())

        #read cassys csyx config file
        self.version = self.csyx['Site']['Version']
        self.orientation = self.csyx['Site']['Orientation_and_Shading']
        self.array_type = self.csyx['Site']['Orientation_and_Shading']['@ArrayType']

        self.nameplateDCkW = float(self.csyx['Site']['System']['SystemDC'])
        self.numSubArray = int(self.csyx['Site']['System']['@TotalArrays'])


    #Run CASSYS with live stdout
    #https://www.endpoint.com/blog/2015/01/28/getting-realtime-output-using-python
    def run(self,output_file = None, verbose = False):
        # write csyx_file in case some modificatin where made to dict

        '''
        pwd
        self = cassysObj(r'pycassys\tests\Bifacial_Model_1.csyx',r'pycassys\tests\ACA_climate_file.csv')
        verbose = True
        '''
        self.write_csyx()

        # create temp file if none specified
        if output_file is None:
            output_file = os.path.join(self.output_path, 'cassys_temp_out.csv') # create temp file_path
        if os.path.exists(output_file):
            os.remove(output_file)
        expOutput_file = output_file # os.path.join(self.output_path,output_file)
        arguments = [self.config_file,self.climate_file,expOutput_file]
        command = [self.cassys_exe]

        command.extend(arguments)
        """Launches 'command' windowless and waits until finished"""
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        process = subprocess.Popen(command, stdout=subprocess.PIPE,encoding='utf8', startupinfo=startupinfo)

        # prepare to check iterrows

        if os.path.exists(self.errorlog): os.remove(self.errorlog)

        while True:
            output = process.stdout.readline()
            if output == '' and process.poll() is not None:
                break
            if output:
                if verbose: print(output.strip())

        rc = process.poll()

        # check for errot log
        if os.path.exists(expOutput_file):
            df = pd.read_csv(expOutput_file)
            if len(df) <8760:
                raise Exception('error running CASSYS: Check error log')
            else:
                return df
        else:
            if os.path.exists(self.errorlog):
                with open(self.errorlog) as f:
                    err = f.read()
                # TODO replace with raise error
                raise Exception('error running CASSYS: {} \n'.format(err))
                df = pd.DataFrame()
            else:
                raise Exception('error running CASSYS: {} \n'.format(err))


    #Run CASSYS via DLL not workign
    #https://www.endpoint.com/blog/2015/01/28/getting-realtime-output-using-python
    def dll_run(self,output_file):
        expOutput_file = os.path.join(self.output_path,output_file)
        args = [self.config_file,self.climate_file, expOutput_file]
        mycassys = CASSYSClass()
        mycassys.Main(args)

        df = pd.read_csv(expOutput_file)

        return df

    def write_csyx(self):
        with open(self.config_file,'w') as fd:
            fd.write(xmltodict.unparse(self.csyx, pretty=True))

            # print("{} file writen with new parameters".format(self.config_file))
