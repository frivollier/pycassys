"""
Canadian Solar SDTools CASSYS module
Logging in Power Shell Get-content  c:\pyxll\logs\pyxll.20180802.log -Tail 0 -Wait


# START HERE if you are debuging

run hydrogen in SDTools env

%load_ext autoreload
%autoreload 2
%gui qt

from plotly.offline import download_plotlyjs, init_notebook_mode, plot, iplot
init_notebook_mode(connected=True)


"""
#%% imports
import sys, os, inspect
import subprocess
import shutil
from datetime import datetime
from uuid import uuid4
#PYXLL imports
from pyxll import xl_func, xl_macro, xlcAlert, xl_menu, xl_app, xlfCaller, get_active_object
from pyxll_utils.pandastypes import _dataframe_to_var, _series_to_var, _series_to_var_transform
from pyxll import xlAsyncReturn, get_type_converter

from threading import Thread
import time
import csv
import tempfile

import pandastypes
import numpy as np
import pandas as pd
import math
import logging
_log = logging.getLogger(__name__)

import utilities

import json

import pycassys
import tempfile
import pvsyst

import pvlib

from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog
from PyQt5.QtGui import QIcon


# QT dialogs
class App(QWidget):

    def __init__(self,):
        super().__init__()
        self.title = 'Dialog'
        self.left = 100
        self.top = 100
        self.width = 640
        self.height = 480
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        # self.show()

    def openFileNameDialog(self, message = 'File Dialog', dir = '', filter = "All Files (*)"):
        options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,message, dir,filter, options=options)

        return  fileName

    def openFileNamesDialog(self):
        options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog
        files, _ = QFileDialog.getOpenFileNames(self,"QFileDialog.getOpenFileNames()", "","All Files (*);;Python Files (*.py)", options=options)
        return  files

    def saveFileDialog(self):
        options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","All Files (*);;Text Files (*.txt)", options=options)
        return fileName

# QT generic dialog to select file
def select_file(message='select a file', dir=r'C:\sdtools\Databases', filter = None):
    app = QApplication(sys.argv)
    ex = App()
    """Select a file via a dialog and return the file name."""
    if dir is None: dir ='./'
    if filter is None: filter = "All files (*);; CASSYS Files (*.csyx)"
    fname = ex.openFileNameDialog(message = message, dir = dir, filter=filter)
    # sys.exit(app.exec_())
    return fname

#%%
@xl_macro ()
def select_climate_file():
    xl = xl_app()
    currentdir = xl.ActiveWorkbook.Path
    climate_file = select_file(message = 'Select TMY3 file', dir = currentdir, filter='CSV (*.csv);; All Files(*)')
    '''
    climate_file = r'C:\sdtools\XLStemplates\MN_TMY3\Australia_-27_ 142_tmy3.csv'
    '''
    if climate_file:
        xl.Range('cassys_tmy3_path').Value = climate_file
        tmy_describe(xl, climate_file) #update al summary fields
        # xlcAlert("Clim file path updated")

# update sumamry field based on TMY data
def tmy_describe(xl, climate_file):
    tmy_data, meta = pvlib.iotools.read_tmy3(climate_file)
    '''
    climate_file = r'C:\sdtools\XLStemplates\MN_TMY3\Australia_-27_ 142_tmy3.csv'
    tmy_data.head()
    tmy_data.columns

    DHI
    GHI
    DryBulb
    type(float(tmy_data['GHI'].sum()/1000))
    float(tmy_data['DryBulb'].mean())
    std = tmy_data['DryBulb'].std()
    std
    tmin = tmy_data['DryBulb'].min()
    tmean = tmy_data['DryBulb'].mean()
    tmin - tmean
    (tmin - tmean)/ std
    # P75 0.675 * std
    # P90 1.282 * std
    # P95 = 1.960
    # P99 2.326 * std
    tmin
    (2.326/1.282 * (tmin - tmean)) + tmean

    std =
    meta['latitude']
    meta['longitude']
    '''
    p95 = 1.960
    p99 = 2.326
    xl.Range('cassys_latlon').Value = '{},{}'.format(meta['latitude'],meta['longitude'])
    xl.Range('cassys_ghi').Value = float(tmy_data['GHI'].sum()/1000)
    xl.Range('cassys_dhi').Value = float(tmy_data['DHI'].sum()/1000)
    tmax = tmy_data['DryBulb'].max()
    tmin = tmy_data['DryBulb'].min()
    tmean = tmy_data['DryBulb'].mean()
    xl.Range('cassys_tamb').Value = float(tmean)
    xl.Range('cassys_tmax').Value = float(tmax)
    xl.Range('cassys_tmin').Value = float(tmin)
    # estimate Textreme assuming TMY lowest values are P95 and extremes are p99
    xl.Range('cassys_tmax_extreme').Value = (p99/p95 * (tmax - tmean)) + tmean
    xl.Range('cassys_tmin_extreme').Value = (p99/p95 * (tmin - tmean)) + tmean


@xl_macro ()
def select_file_path():
    xl = xl_app()
    currentdir = "C:\sdtools\Databases\PAN"
    file = select_file(message = 'Select file path to paste in selected cell', dir = currentdir, filter='All Files(*)')

    if file:
        xl.Selection.Value = file
        # xlcAlert("Clim file path updated")


#%% Main entry point
@xl_macro ("")
def gcr_sweep():
    ''' This is call from t.cassys.xlms RUN SELECTED variant_to_run

    xl.Names
    nms = []
    for n in xl.Names:
        nms.append(n.Name)
    nms
    '''
    xl = xl_app()

    xl.Range('cassys_run_status').Value = 'Initializing...'

    # build dict of field in worksheet
    ws_fields = {}
    for n in xl.Names:
        if n.Name.startswith('cassys_'):
            try:
                ws_fields[n.Name[7:]] = xl.Range(n.Name).Value
            except:
                pass

    # sitename = xl.Range("cassys_sitename").Value
    config_file = os.path.join(r'c:\sdtools\PYscripts\pycassys', 'csyx_templates', ws_fields['csyx'])

    currentdir = xl.ActiveWorkbook.Path
    out_directory = os.path.join(currentdir,"cassys output")
    if not os.path.exists(out_directory): os.makedirs(out_directory)

    #create temporary csyx file for manipulations
    # temp_csyx = os.path.join(tempfile.gettempdir(), 'cassys_temp_conf.csyx') # create temp file_path
    temp_csyx = os.path.join(out_directory, 'cassys_temp_conf.csyx') # create temp file_path
    shutil.copy2(config_file, temp_csyx) # copy to temp file path
    # create temporary output csv file
    # initiate cassys Class
    cassys = pycassys.cassysObj(temp_csyx, ws_fields['tmy3_path'], out_directory)

    # get selected row
    variant_header = [None] * 10
    for i in range(10):
        variant_header[i] =  xl.ActiveWorkbook.Names('cassys_var{}'.format(i)).Value.split('!')[1]

    # prepare variant matrix
    selection = xl.Selection
    variant_to_run = []
    for cell in selection:
        if cell.Address in variant_header:
            variant_to_run.append(variant_header.index(cell.Address))

    if len(variant_to_run) > 0:
        xl.Range('cassys_run_status').Value = 'Ready to run variant(s): {}'.format(variant_to_run)
        print('Ready to run variant(s): {}'.format(variant_to_run))
    else:
        xl.Range('cassys_run_status').Value = ('No variant selected. Please select the varient number cells you want to run ')
        print('No variant selected. Please select the varient number cells you want to run ')
        # return

    # create DF to store results
    df = pd.DataFrame()
    for v in variant_to_run:
        xl.Range('cassys_run_status').Value = 'Running {}'.format(v)
        xl.ScreenUpdating = True
        df = df.append(cassys_run_variant(xl, cassys, ws_fields, v))

    # prepare csv file name
    rid = datetime.now().strftime('%yy%m-%d%H-%M%S') #  + str(uuid4())
    out_csv = os.path.join(out_directory,'{}.csv'.format(rid))
    # save to csv
    df.to_csv(out_csv, index=True)
    xl.Range('cassys_run_status').Value = 'Completed'

    # create new sheet for resutls
    create_result_tab('results_' + rid, df)



    '''
    df['yield'].plot()
    https://www.pyxll.com/_forum/index.php/topic,1234.0.html
    https://github.com/pyxll/develop-excel-london-2018
    https://github.com/pyxll/pyxll-utils/blob/master/pyxll_utils/pandastypes.py

    '''

    # create new pivot graph
    '''
    https://www.thespreadsheetguru.com/blog/2014/9/27/vba-guide-excel-pivot-tables
    '''

    return True


def create_result_tab(name, df):
    xl = xl_app()
    '''
    name = 'test'
    df = pd.DataFrame()
    result_ws = xl.ActiveSheet
    org = result_ws.Range('A30:AF50')
    '''
    # past resutls in sheet
    xl.ScreenUpdating = False
    filename = os.path.join(r'c:\sdtools\XLStemplates','o.cassys.xlsm')
    target_wb = xl.ActiveWorkbook
    source_wb = xl.Workbooks.Open(filename)
    result_ws = source_wb.Worksheets['o.cassys']
    result_ws.Copy(After=target_wb.Worksheets(target_wb.Sheets.Count))
    xl.DisplayAlerts = False
    source_wb.Close()
    xl.DisplayAlerts = True
    xl.ScreenUpdating = True

    result_ws = target_wb.Worksheets['o.cassys']
    result_ws.Name = name # reasign to new sheet
    df.reset_index(level=0, inplace=True)
    org = result_ws.Range('$A$30')
    org = result_ws.Range(org, org.Resize(df.shape[0]+1, df.shape[1])) # Resize
    org.Value = _dataframe_to_var(df)

    for pt in result_ws.PivotTables():
        pt.SourceData = org.Address
        pt.RefreshTable

    '''
    target_wb.Name
    source_wb.Name
    target_wb.Sheets.Count
    pt = result_ws.PivotTables()[1]
    pt.SourceData = org.Address
    org.Address
    pt.RefreshTable
    '''




# update csyx dict from ws_fields
def update_csyx(csyx, ws_fields):
    '''
    update csyx XML file from dict

    '''
    # change locations
    lat = float(ws_fields['latlon'].split(",")[0])
    lon = float(ws_fields['latlon'].split(",")[1])
    csyx['Site']['SiteDef']['Latitude'] = lat
    csyx['Site']['SiteDef']['Longitude'] = lon
    csyx['Site']['SiteDef']['Altitude'] = utilities.get_altidude(lat, lon)
    csyx['Site']['SiteDef']['TimeZone'] = utilities.get_UTC_offset(lat, lon)
    csyx['Site']['SiteDef']['RefMer']  = csyx['Site']['SiteDef']['TimeZone']  * 15
    csyx['Site']['SiteDef']['UseLocTime'] = True
    # set albedo
    csyx['Site']['SiteDef']['Albedo']['Jan'] = ws_fields['albedo_jan']
    csyx['Site']['SiteDef']['Albedo']['Feb'] = ws_fields['albedo_feb']
    csyx['Site']['SiteDef']['Albedo']['Mar'] = ws_fields['albedo_mar']
    csyx['Site']['SiteDef']['Albedo']['Apr'] = ws_fields['albedo_apr']
    csyx['Site']['SiteDef']['Albedo']['May'] = ws_fields['albedo_may']
    csyx['Site']['SiteDef']['Albedo']['Jun'] = ws_fields['albedo_jun']
    csyx['Site']['SiteDef']['Albedo']['Jul'] = ws_fields['albedo_jul']
    csyx['Site']['SiteDef']['Albedo']['Aug'] = ws_fields['albedo_aug']
    csyx['Site']['SiteDef']['Albedo']['Sep'] = ws_fields['albedo_sep']
    csyx['Site']['SiteDef']['Albedo']['Oct'] = ws_fields['albedo_oct']
    csyx['Site']['SiteDef']['Albedo']['Nov'] = ws_fields['albedo_nov']
    csyx['Site']['SiteDef']['Albedo']['Dec'] = ws_fields['albedo_dec']

    # set soiling
    csyx['Site']['SoilingLosses']['Jan'] = ws_fields['soiling_jan']
    csyx['Site']['SoilingLosses']['Feb'] = ws_fields['soiling_feb']
    csyx['Site']['SoilingLosses']['Mar'] = ws_fields['soiling_mar']
    csyx['Site']['SoilingLosses']['Apr'] = ws_fields['soiling_apr']
    csyx['Site']['SoilingLosses']['May'] = ws_fields['soiling_may']
    csyx['Site']['SoilingLosses']['Jun'] = ws_fields['soiling_jun']
    csyx['Site']['SoilingLosses']['Jul'] = ws_fields['soiling_jul']
    csyx['Site']['SoilingLosses']['Aug'] = ws_fields['soiling_aug']
    csyx['Site']['SoilingLosses']['Sep'] = ws_fields['soiling_sep']
    csyx['Site']['SoilingLosses']['Oct'] = ws_fields['soiling_oct']
    csyx['Site']['SoilingLosses']['Nov'] = ws_fields['soiling_nov']
    csyx['Site']['SoilingLosses']['Dec'] = ws_fields['soiling_dec']

    if ws_fields['pan']:
        module = read_pan(ws_fields['pan'])
        csyx = _update_module_parameters(csyx, module)

    # set String Length
    if ws_fields['string_length']:
        csyx['Site']['System']['SubArray1']['NumStrings'] = int(ws_fields['string_length'])

    # DC Ohmic loss
    if ws_fields['dc_ohmic_loss']:
        csyx['Site']['System']['SubArray1']['PVModule']['LossFraction'] = float(ws_fields['dc_ohmic_loss'])

    # AC Ohmic loss
    if ws_fields['ac_loss']:
        csyx['Site']['System']['SubArray1']['Inverter']['LossFraction'] = float(ws_fields['ac_loss'])

    # Transformer  Ohmic loss
    transfo_pnom = csyx['Site']['System']['SystemAC']
    csyx['Site']['Transformer']['PNomTrf'] = transfo_pnom

    if ws_fields['iron_loss']:
        csyx['Site']['Transformer']['PIronLossTrf'] = float(ws_fields['iron_loss']) * float(transfo_pnom) # in kW

    if ws_fields['transfo_ohmic_loss']:
        csyx['Site']['Transformer']['PResLssTrf'] = float(ws_fields['transfo_ohmic_loss']) * float(transfo_pnom) # in kW

    # total losses FULL LOAD
    csyx['Site']['Transformer']['PFullLoadLss'] = (float(ws_fields['transfo_ohmic_loss'])+ float(ws_fields['iron_loss'])) * float(transfo_pnom)

    # AC capacity at ACCapSTC RESET to 0 as it is not used by CASYS Engine
    csyx['Site']['Transformer']['ACCapSTC'] = 0 # not used by engine

    # cassys_night_disconnect
    if ws_fields['night_disconnect'] == 'Yes':
        csyx['Site']['Transformer']['NightlyDisconnect'] = True
    else:
        csyx['Site']['Transformer']['NightlyDisconnect'] = False

    # Bifacial parameters
    # cassys_night_disconnect
    if ws_fields['bifacial'] == 'Yes':
        csyx['Site']['Bifacial']['UseBifacialModel'] = True
    else:
        csyx['Site']['Bifacial']['UseBifacialModel'] = False

    # cassys_bifi_ground_clearance
    if ws_fields['bifi_ground_clearance']:
        csyx['Site']['Bifacial']['GroundClearance'] = float(ws_fields['bifi_ground_clearance'])

    # cassys_rear_blocking_factor
    if ws_fields['rear_blocking_factor']:
        csyx['Site']['Bifacial']['StructBlockingFactor'] = float(ws_fields['rear_blocking_factor'])

    # cassys_bifi_transmision_factor
    if ws_fields['bifi_transmision_factor']:
        csyx['Site']['Bifacial']['PanelTransFactor'] = float(ws_fields['bifi_transmision_factor'])

    # cassys_bifaciality
    if ws_fields['bifaciality']:
        csyx['Site']['Bifacial']['BifacialityFactor'] = float(ws_fields['bifaciality'])

    # set bifacial albedo
    csyx['Site']['Bifacial']['BifAlbedo']['Jan'] = ws_fields['albedo_jan']
    csyx['Site']['Bifacial']['BifAlbedo']['Feb'] = ws_fields['albedo_feb']
    csyx['Site']['Bifacial']['BifAlbedo']['Mar'] = ws_fields['albedo_mar']
    csyx['Site']['Bifacial']['BifAlbedo']['Apr'] = ws_fields['albedo_apr']
    csyx['Site']['Bifacial']['BifAlbedo']['May'] = ws_fields['albedo_may']
    csyx['Site']['Bifacial']['BifAlbedo']['Jun'] = ws_fields['albedo_jun']
    csyx['Site']['Bifacial']['BifAlbedo']['Jul'] = ws_fields['albedo_jul']
    csyx['Site']['Bifacial']['BifAlbedo']['Aug'] = ws_fields['albedo_aug']
    csyx['Site']['Bifacial']['BifAlbedo']['Sep'] = ws_fields['albedo_sep']
    csyx['Site']['Bifacial']['BifAlbedo']['Oct'] = ws_fields['albedo_oct']
    csyx['Site']['Bifacial']['BifAlbedo']['Nov'] = ws_fields['albedo_nov']
    csyx['Site']['Bifacial']['BifAlbedo']['Dec'] = ws_fields['albedo_dec']


    # set orientation
    if ws_fields['orientation'] == 'Unlimited Rows':
        #type = 'Fixed Tilted Plane'
        csyx['Site']['Orientation_and_Shading']['@ArrayType'] = ws_fields['orientation']

        # adjust orientation and shading
        csyx['Site']['Orientation_and_Shading']['CollBandWidth'] = float(ws_fields['chord'])
        csyx['Site']['Orientation_and_Shading']['TopInactive'] = 0
        csyx['Site']['Orientation_and_Shading']['BottomInactive'] = 0
        if csyx['Site']['SiteDef']['Latitude'] >= 0:
            csyx['Site']['Orientation_and_Shading']['Azimuth'] = 0
        else:
            csyx['Site']['Orientation_and_Shading']['Azimuth'] = 180
        csyx['Site']['Orientation_and_Shading']['RowsBlock'] = 50
        csyx['Site']['Orientation_and_Shading']['UseCellVal'] = 'True'
        csyx['Site']['Orientation_and_Shading']['StrInWid'] = 2
        csyx['Site']['Orientation_and_Shading']['CellSize'] = 15.6
        csyx['Site']['Orientation_and_Shading']['WidOfStr'] = float(ws_fields['chord'])/csyx['Site']['Orientation_and_Shading']['StrInWid']
        csyx['Site']['Orientation_and_Shading']['DefineHorizonProfile'] = 'False'

    elif ws_fields['orientation'] == 'Single Axis Horizontal Tracking (N-S)':
        # set Orientation to SAT
        csyx['Site']['Orientation_and_Shading']['@ArrayType'] = ws_fields['orientation']
        # adjust orientation and shading
        csyx['Site']['Orientation_and_Shading']['WActiveSAST'] = float(ws_fields['chord'])
        csyx['Site']['Orientation_and_Shading']['AxisTiltSAST'] = 0

        if csyx['Site']['SiteDef']['Latitude'] >= 0:
            csyx['Site']['Orientation_and_Shading']['AxisAzimuthSAST'] = 0
        else:
            csyx['Site']['Orientation_and_Shading']['AxisAzimuthSAST'] = 0
        csyx['Site']['Orientation_and_Shading']['RotationMaxSAST'] = 60
        csyx['Site']['Orientation_and_Shading']['RowsBlockSAST'] = 50
        csyx['Site']['Orientation_and_Shading']['BacktrackOptSAST'] = 'True'
        csyx['Site']['Orientation_and_Shading']['UseCellValSAST'] = 'False'
        csyx['Site']['Orientation_and_Shading']['DefineHorizonProfile'] = 'False'

    # Set ThermalLosses
    if ws_fields['uc']:
        csyx['Site']['Losses']['ThermalLosses']['ConsHLF'] = float(ws_fields['uc'])

    # quality loss typicaly negative -0.3% = -0.003
    if ws_fields['quality_loss']:
        csyx['Site']['Losses']['ModuleQualityLosses']['PowerLoss'] = float(ws_fields['quality_loss'])

    # LID  1.5% = -0.015
    if ws_fields['lid']:
        csyx['Site']['Losses']['ModuleQualityLosses']['ModuleLID'] = float(ws_fields['lid'])

    # Agging loss
    if ws_fields['ageing']:
        csyx['Site']['Losses']['ModuleQualityLosses']['ModuleAgeing'] = float(ws_fields['ageing'])

    # Mismatch loss
    if ws_fields['mismatch']:
        csyx['Site']['Losses']['ModuleMismatchLosses']['PowerLoss'] = float(ws_fields['mismatch'])


    if ws_fields['orientation'] == 'Unlimited Rows':
        csyx['Site']['Orientation_and_Shading']['Pitch'] = float(ws_fields['pitch'])
        a = math.sin(math.radians(float(ws_fields['tilt'])))* float(ws_fields['chord'])
        shadinglimit =  math.degrees(math.atan(a/float(ws_fields['pitch'])))
        csyx['Site']['Orientation_and_Shading']['ShadingLimit'] = shadinglimit
        csyx['Site']['Orientation_and_Shading']['PlaneTilt'] = float(ws_fields['tilt'])
    else:
        csyx['Site']['Orientation_and_Shading']['PitchSAST'] = float(ws_fields['pitch'])

# return col of named range
def named_range_col(named_range):
    xl = xl_app()
    return xl.Range(named_range).Address.split('$')[1]

# return row of named range
def named_range_row(named_range):
    xl = xl_app()
    return xl.Range(named_range).Address.split('$')[2]

# return value of based on variant (lookup col row)
def get_variant_value(variant, value):
    xl = xl_app()
    row = named_range_row(value)
    return xl.Range('${}${}'.format(named_range_col(variant),row)).Value

# run all sim for one variant
def cassys_run_variant(xl, cassys, ws_fields, variant):
    ''' Run Cassys based on Variant (column) paramters
    parameters:
        cassys :cassy Class instance
        variant: int of column ID

    '''
    # get variant field name and column
    var = 'cassys_var{}'.format(variant)
    var_col = named_range_col(var) #column in excel

    # get variant name
    var_name = get_variant_value(var, 'cassys_variant_name')
    print('CASSYS running variant: [{}]'.format(var_name))

    # set orientation
    ws_fields['orientation'] = get_variant_value(var, 'cassys_orientation')

    ws_fields['arrays'] = int(cassys.csyx['Site']['System']['@TotalArrays'])

    # set module parameters
    ws_fields['pan'] = get_variant_value(var, 'cassys_pan')
    # set string Length (We assum that there is only one sub array)
    ws_fields['string_length'] = get_variant_value(var, 'cassys_string_length')
    # DC Ohmic loss
    ws_fields['dc_ohmic_loss']  = get_variant_value(var, 'cassys_dc_ohmic_loss')
    # AC Ohmic loss
    ws_fields['ac_loss']  = get_variant_value(var, 'cassys_ac_loss')
    ws_fields['iron_loss']  = get_variant_value(var, 'cassys_iron_loss') # percent
    ws_fields['transfo_ohmic_loss']  = get_variant_value(var, 'cassys_transo_ohmic_loss') # percent
    # cassys_night_disconnect
    ws_fields['night_disconnect']  = get_variant_value(var, 'cassys_night_disconnect') # percent
    # Bifacial parameters
    # cassys_night_disconnect
    ws_fields['bifacial']  = get_variant_value(var, 'cassys_bifacial') # Yes or No
    # cassys_bifi_ground_clearance
    ws_fields['bifi_ground_clearance']  = get_variant_value(var, 'cassys_bifi_ground_clearance') # m
    # cassys_rear_blocking_factor
    ws_fields['rear_blocking_factor']  = get_variant_value(var, 'cassys_rear_blocking_factor')
    # cassys_bifi_transmision_factor
    ws_fields['bifi_transmision_factor']  = get_variant_value(var, 'cassys_bifi_transmision_factor')
    # cassys_bifaciality
    ws_fields['bifaciality']  = get_variant_value(var, 'cassys_bifaciality')
    # get chord
    ws_fields['chord'] = get_variant_value(var, 'cassys_chord') or 2
    # Set Losses
    ws_fields['uc'] = get_variant_value(var, 'cassys_uc')
    # quality loss typicaly negative -0.3% = -0.003
    ws_fields['quality_loss'] = get_variant_value(var, 'cassys_quality_loss')
    # LID  1.5% = -0.015
    ws_fields['lid'] = get_variant_value(var, 'cassys_lid')
    # Agging loss
    ws_fields['ageing']  = get_variant_value(var, 'cassys_ageing')
    # Mismatch loss
    ws_fields['mismatch']  = get_variant_value(var, 'cassys_mismatch')

    # get ILRs
    ilr_min = get_variant_value(var, 'cassys_ilr_min')
    ilr_max = get_variant_value(var, 'cassys_ilr_max')
    ilr_steps = get_variant_value(var, 'cassys_ilr_steps')
    if ilr_steps is None: ilr_steps = 2

    if (ilr_min is not None) and (ilr_max is not None):
            if ilr_steps > 2:
                ilrs = np.linspace(ilr_min,ilr_max, num = int(ilr_steps))
            else:
                ilrs = np.linspace(ilr_min,ilr_max, num = 2)
    else:
        if max(ilr_max or 0, ilr_min or 0) > 0:
                ilrs = [max(ilr_max or 0, ilr_min or 0)]
        else:
            ilrs = [1.25] # if not values run at 1.25

    print('ilrs = {}'.format(ilrs))

    # set GCRs
    gcr_min = get_variant_value(var, 'cassys_gcr_min')
    gcr_max = get_variant_value(var, 'cassys_gcr_max')
    gcr_steps = get_variant_value(var, 'cassys_gcr_steps')
    if gcr_steps is None: gcr_steps = 2

    if (gcr_min is not None) and (gcr_max is not None):
            if gcr_steps > 2:
                gcrs = np.linspace(gcr_min,gcr_max, num = int(gcr_steps))
            else:
                gcrs = np.linspace(gcr_min,gcr_max, num = 2)
    else:
        if max(gcr_max or 0, gcr_min or 0) > 0:
                gcrs = [max(gcr_max or 0, gcr_min or 0)]
        else:
            gcrs = [0.3] # if not values run at 0.3

    print('gcrs = {}'.format(gcrs))


    # set Tilts
    if ws_fields['orientation'] == 'Unlimited Rows':
        tilt_min = get_variant_value(var, 'cassys_tilt_min')
        tilt_max = get_variant_value(var, 'cassys_tilt_max')
        tilt_steps = get_variant_value(var, 'cassys_tilt_steps')
        if tilt_steps is None: tilt_steps = 2

        if (tilt_min is not None) and (tilt_max is not None):
                if tilt_steps > 2:
                    tilts = np.linspace(tilt_min,tilt_max, num = int(tilt_steps))
                else:
                    tilts = np.linspace(tilt_min,tilt_max, num = 2)
        else:
            if max(tilt_max or 0, tilt_min or 0) > 0:
                    tilts = [max(tilt_max or 0, tilt_min or 0)]
            else:
                tilts = [15] # if not values run at 15

        print('tilts = {}'.format(tilts))

    else:
        tilts = [0] # set to 0 for SAT
        print('tilts = {} - Single Axis Tracker'.format(tilts))

    runs = len(gcrs) * len(ilrs) * len(tilts)

    # build empty DataFrame of all runs
    variant_runs = pd.DataFrame()
    for idg, gcr in np.ndenumerate(gcrs):
        for idt, tilt in np.ndenumerate(tilts):
            for idi, ilr in np.ndenumerate(ilrs):
                r = {}
                id = '{}_{}_{}_{}'.format(var_name, gcr, tilt, ilr)
                r['variant'] = var_name
                # variant parameters
                r['Orientation'] = ws_fields['orientation']
                r['Chord']  = float(ws_fields['chord'])

                # variant run values
                r['GCR'] = gcr
                r['Tilt'] = tilt
                r['ILR'] = ilr

                # prepare output_values
                r['GHI'] = 0.0
                r['DHI'] = 0.0
                r['GTI'] = 0.0
                r['Transpostion Boost (%)'] = 0.0
                r['Far Shadings: irradiance loss'] = 0.0
                r['Near Shadings: irradiance loss'] = 0.0
                r['IAM factor on global'] = 0.0
                r['Soiling loss factor'] = 0.0
                r['Effective irradiation on collectors'] = 0.0
                r['PV loss due to irradiance level'] = 0.0
                r['PV loss due to temperature'] = 0.0
                r['Module quality loss'] = 0.0
                r['LID - Light induced degradation'] = 0.0
                r['Mismatch loss, modules and strings'] = 0.0
                r['Ohmic wiring loss'] = 0.0
                r['Array virtual energy at MPP'] = 0.0
                r['Inverter Loss during operation (efficiency)'] = 0.0
                r['Inverter Loss over nominal inv. power'] = 0.0
                r['Night consumption'] = 0.0
                r['Auxiliaries (fans, other)'] = 0.0
                r['AC ohmic loss'] = 0.0
                r['External transfo loss'] = 0.0
                r['Active Energy injected into grid'] = 0.0
                r['Yield'] = 0.0



                variant_runs = variant_runs.append(pd.DataFrame(r, index=[id]), sort=False)

    # iterate through DF an run CASSYS
    for index, run in variant_runs.iterrows():
        # adjust Pitch and tilts
        gcr = run['GCR']
        tilt = run['Tilt']
        chord = run['Chord']
        pitch = chord / gcr

        ws_fields['tilt'] = tilt
        ws_fields['chord'] = chord
        ws_fields['pitch'] = pitch

        # push all ws_fields dict chnages to csyx
        update_csyx(cassys.csyx ,ws_fields)

        # adjust ILR
        ilr = run['ILR']
        adjust_ilr(cassys.csyx, ilr)

        # run
        print('running: {}'.format(index))
        # RUN CASSYS csyx will be writed ot xml autoamticaly
        res = cassys.run(verbose = False)

        # prepare outputs
        kWp = float(cassys.csyx['Site']['System']['SystemDC'])

        variant_runs.loc[index, 'yield'] = res['Energy Injected into Grid (kWh)'].sum()/kWp

        ghi = res['Horizontal Global Irradiance (W/m2)'].sum()
        variant_runs.loc[index, 'GHI'] = ghi/1000
        variant_runs.loc[index, 'DHI'] = res['Horizontal diffuse irradiance (W/m2)'].sum()/1000

        gti = res['Global Irradiance in Array Plane (W/m2)'].sum()
        variant_runs.loc[index, 'GTI'] = gti/1000

        variant_runs.loc[index, 'Transpostion Boost (%)'] = (gti/ghi)-1

        total_far_shading = res['Horizon Shading Loss for Global (W/m2)'].sum()
        variant_runs.loc[index, 'Far Shadings: irradiance loss'] = total_far_shading/gti

        total_near_shading = res['Near Shading Loss for Global (W/m2)'].sum()
        variant_runs.loc[index, 'Near Shadings: irradiance loss'] = total_near_shading/gti

        iam_loss = res['Incidence Loss for Global (W/m2)'].sum()
        variant_runs.loc[index, 'IAM factor on global'] = iam_loss/gti

        soiling_loss = res['Soiling Loss (W/m2)'].sum()
        variant_runs.loc[index, 'Soiling loss factor'] = soiling_loss/gti

        # ADD BIFI HERE

        variant_runs.loc[index, 'Effective irradiation on collectors'] = res['Effective Global Irradiance in Array Plane (W/m2)'].sum()/1000
        variant_runs.loc[index, 'PV loss due to irradiance level'] = 0.0
        variant_runs.loc[index, 'PV loss due to temperature'] = 0.0
        variant_runs.loc[index, 'Module quality loss'] = 0.0
        variant_runs.loc[index, 'LID - Light induced degradation'] = 0.0
        variant_runs.loc[index, 'Mismatch loss, modules and strings'] = 0.0
        variant_runs.loc[index, 'Ohmic wiring loss'] = 0.0
        variant_runs.loc[index, 'Array virtual energy at MPP'] = 0.0
        variant_runs.loc[index, 'Inverter Loss during operation (efficiency)'] = 0.0
        variant_runs.loc[index, 'Inverter Loss over nominal inv. power'] = 0.0
        variant_runs.loc[index, 'Night consumption'] = 0.0
        variant_runs.loc[index, 'Auxiliaries (fans, other)'] = 0.0
        variant_runs.loc[index, 'AC ohmic loss'] = 0.0
        variant_runs.loc[index, 'External transfo loss'] = 0.0
        variant_runs.loc[index, 'Active Energy injected into grid'] = 0.0
        variant_runs.loc[index, 'Yield'] = 0.0

        # add ghi

        # add clipping losses


    return variant_runs


# adjust csyx ilr
def adjust_ilr(csyx, ilr = 1.3):
    '''
    csyx = 'one_block_monofacial.csyx'
    config_file = os.path.join(r'c:\sdtools\PYscripts\pycassys', 'csyx_templates', csyx)
    climate_file = 'C:/sdtools/XLStemplates/MN_TMY3/Test Project_27.935729_-15.571356_tmy3.csv'

    temp_csyx = os.path.join(tempfile.gettempdir(), 'cassys_temp_conf.csyx') # create temp file_path
    shutil.copy2(config_file, temp_csyx) # copy to temp file path
    # create temporary output csv file
    cassys = pycassys.cassysObj(temp_csyx,climate_file,'.')

    ilr = 2.0
    '''

    ac_nameplate = float(csyx['Site']['System']['SystemAC']) * 1000
    target_dc = ac_nameplate * ilr

    p_nom = float(csyx['Site']['System']['SubArray1']['PVModule']['Pnom'])
    modules_in_string = int(csyx['Site']['System']['SubArray1']['PVModule']['ModulesInString'])
    strings = float(csyx['Site']['System']['SubArray1']['PVModule']['NumStrings'])
    string_power = p_nom * modules_in_string
    strings = int(target_dc/string_power)


    csyx['Site']['System']['SubArray1']['PVModule']['NumStrings'] = strings
    # adjust Global DC wire resistance
    '''
    Power loss = I^2*R
    '''
    dc_loss = csyx['Site']['System']['SubArray1']['PVModule']['LossFraction']
    i = csyx['Site']['System']['SubArray1']['PVModule']['Impp'] * strings
    r = (p_nom*dc_loss)/(i**2) # in Ohm
    csyx['Site']['System']['SubArray1']['PVModule']['GlobWireResist'] = r * 1000  # in mOhm

    #change other related values
    csyx['Site']['System']['SubArray1']['PVModule']['NumModules'] = int(modules_in_string) * int(strings)
    csyx['Site']['System']['SystemDC'] = (int(modules_in_string) * int(strings) * int(p_nom) /1000)


    '''
    @debug Tests
    target_dc
    strings
    print("New String Quantity = {}".format(strings))
    print("Module (W) = {}".format(p_nom))
    print("Module per String = {}".format(modules_in_string))
    print("CASSYS Version = {}".format(csyx['Site']['Version']))
    print("Number of Array(s) = {}".format(csyx['Site']['System']['@TotalArrays']))
    print("DC Nameplate (kW) = {}".format(float(csyx['Site']['System']['SystemDC'])))
    print("Current String quantity = {}".format(strings))
    print("New DC Nameplate (kW) = {}".format(float(csyx['Site']['System']['SystemDC'])))
    print("New DC Nameplate (kW) = {}".format((int(ModulesInString) * int(strings) * int(Pnom) /1000)))

    '''
    return csyx


# read pan file to dict
def read_pan(pan_path):
    ''''
    pan_path = r'C:/SDTools/Databases/PAN/HiKu Poly/CS3W-P/CS3W-405P_MIX_CSIHE_PRE_V6_77_1500V_2018Q4.PAN'
    pvsyst.pan_to_dict(pan_path)
    '''
    pan = pvsyst.pan_to_cassys_param(pan_path)
    return pan

@xl_func("string pan_path, string paramater: string")
def get_pan_parameter(pan_path, paramater):
    pan = pvsyst.pan_to_dict(pan_path)
    return pan[paramater]


def _update_module_parameters(csyx, module):
    arrays = int(csyx['Site']['System']['@TotalArrays'])
    for i  in range(arrays):
        '''
        i = 0
        '''
        a = 'SubArray{}'.format(i+1)
        csyx['Site']['System'][a]['PVModule'].update(module)

    return csyx


@xl_macro ()
def getMODIS_CASSYS(lat, lon, name):
    xl = xl_app()
    currentdir = xl.ActiveWorkbook.Path
    climate_file = select_file(message = 'Select TMY3 file', dir = currentdir, filter='CSV (*.csv);; All Files(*)')

    if climate_file:
        xl.Range('cassys_tmy3_path').Value = climate_file
        # xlcAlert("Clim file path updated")
