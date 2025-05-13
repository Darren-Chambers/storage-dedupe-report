'''
Report Storage Metrics from data gathered by DataGather.py in JSON format
Reports are created in Excel format

Report Tabs
    Data Centre View
    Port View
    Host Group View
    LDEV / LUN View
    OS View
    Pool View
Copyright Hitachi Vantara 2025
'''
import argparse
import json
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

from hiraid.raidcom import Raidcom

import xlreport as xlReport

PortChart = ''



#LdevLun Report Variables
class LdevLun():
    def __init__(self):
        self.location              = ''
        self.serialnumber          = ''
        self.arrayModel            = ''
        self.arrayNickname         = ''
        self.port                  = ''
        self.gid                   = ''
        self.hostname              = ''
        self.hostmode              = ''
        self.loggedin              = False
        self.wwn                   = []
        self.lunNumber             = ''
        self.ldevNumberDec         = ''
        self.ldevNumberHex         = ''
        self.ldevLabel             = ''
        self.volAttr               = ''
        self.pool                  = ''
        self.LdevCapacityAlloc_GB  = 0
        self.LdevCapacityUsed_GB   = 0
        self.LdevCapacityUsed_PC   = 0
        self.rsgid                 = ''
        self.virtualSerial         = ''
        self.vsmName               = ''
        self.virtualLdevNumberDec  = ''
        self.virtualLdevNumberHex  = ''

#Port Report Variables
class port():
    def __init__(self):
        self.location             = ''
        self.serialnumber         = ''
        self.arrayModel           = ''
        self.arrayNickname        = ''
        self.port                 = ''
        self.host_groups          = ''
        self.lunCount             = ''
        self.ldevCount            = ''
        self.LdevCapacityAlloc_GB = ''
        self.LdevCapacityUsed_GB  = ''
        self.LdevCapacityUsed_PC  = ''

#HostGroup Report Variables
class hostGroup():
    def __init__(self) -> None:
        self.location             = ''
        self.serialnumber         = ''
        self.arrayModel           = ''
        self.arrayNickname        = ''
        self.port                 = ''
        self.gid                  = ''
        self.hostname             = ''
        self.hostmode             = ''
        self.loggedin             = False
        self.wwn                  = [] 
        self.lunCount             = 0
        self.LdevCapacityAlloc_GB = 0
        self.LdevCapacityUsed_GB  = 0
        self.LdevCapacityUsed_PC  = 0  
        self.rsgid                = []
        self.virtualSerial        = []
        self.vsmName              = []    

#OsType Report Variables
class hostMode():
    def __init__(self) -> None:
        self.location                 = ''
        self.serialnumber             = ''
        self.arrayModel               = ''
        self.arrayNickname            = ''
        self.OsName                   = ''
        self.OsCount                  = 0
        self.lunCount                 = 0
        self.ldevCount                = 0
        self.LdevCapacityAlloc_GB     = 0
        self.LdevcapacityUsed_GB      = 0
        self.LdevcapacityUsed_PC      = 0 

#Pool Report Variables
class pool():
    def __init__(self) -> None:
        self.location            = ''
        self.serialnumber        = ''
        self.arrayModel          = ''
        self.arrayNickname       = ''
        self.poolNumber          = ''      # PID
        self.poolName            = ''      # POOL_NAME
        self.poolType            = ''      # PT
        self.poolCapacity_TB     = 0       # ACT_TP(MB)
        self.poolUsed_TB         = 0       # Capacity(MB) - Available(MB)
        self.poolUsed_PC         = 0       # U(%)
        self.poolFree_TB         = 0       # Available(MB)
        self.poolSubscritpion_TB = 0       # (MappedTB + UnmappedTB) 
        self.poolSubscritpion_PC = 0       # (MappedTB + UnmappedTB) / CapacityTB * 100
        self.poolLdevs           = 0       # LCNT
        self.poolMapped_TB       = 0       # Calculated
        self.poolUnmapped_TB     = 0       # Calculated
        self.poolTotalEfficiecy  = 0       # TOTAL_EFF_R
        self.poolDataReduction   = 0       # TLS_R
        self.poolSoftwareSaving  = 0       # PLS_R
        self.poolCompression     = 0       # PLS_CMP_R
        self.poolDeduplication   = 0       # PLS_DDP_R
        self.poolPatternMatch    = 0       # PLS_RECLAIM_R
        self.poolProvision_PC    = 0       # PROVISIONING_EFF
        self.poolThreshWarn      = 0       # W(%)
        self.poolThreshDeplete   = 0       # H(%)


SerialNumber2Name = {
    "58302": "NP3G1K02O",
    "58303": "PP2G1K05O",
    "350331": "EP3G1K01O",
    "560023": "EP3G56001O",
    "358943": "EP3G1K02O",
    "358372": "PH2G1K05O",
    "540231": "EP3G56002O",
    "540234": "EP3G56003O",
    "560024": "EP2G56002O",
    "560025": "EP2G56001O",
    "540240": "HZBG56001O",
    "560031": "HZAG56001O"
}
SerialNumber2Site = {
    "58302": "PB3",
    "58303": "PB2",
    "350331": "PB3",
    "358372": "HZ",
    "560023": "PB3",
    "540231": "PB3",
    "540234": "PB3",
    "560024": "PB2",
    "560025": "PB2",
    "540240": "HZ",
    "560031": "HZ"
}


def createdir(directory):
    if not os.path.exists(directory):
        os.mkdir(directory)
        

def configlog(scriptname,logdir,logname,basedir=os.getcwd()):
    global log
    try:
        separator = ('/','\\')[os.name == 'nt']
        cwd = basedir
        createdir('{}{}{}'.format(cwd,separator,logdir))
        logfile = '{}{}{}{}{}'.format(cwd, separator, logdir, separator, logname)
        logger = logging.getLogger(scriptname)
        logger.setLevel(logging.DEBUG)
        fh = logging.FileHandler(logfile)
        fh.setLevel(logging.DEBUG)
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s : %(name)s : %(levelname)s : %(message)s')
        fh.setFormatter(formatter)
        ch.setFormatter(formatter)
        # Add handlers to the logger
        logger.addHandler(fh)
        logger.addHandler(ch)
        return logger

    except Exception as e:
        raise Exception('Unable to configure logger > {}'.format(str(e))) 


def CreateDataCentreView(arrayJson, newWorkbook):
    # Create Port Bar Charts
    '''
    newWorkbook.worksheet = newWorkbook.workbook.get_worksheet_by_name("Port View")
    # get all the strings in cells
    sharedStrings = sorted(newWorkbook.worksheet.str_table.string_table, key=newWorkbook.worksheet.str_table.string_table.get)
    titleRow = -1
    titleDict = {}
    for oneRow in newWorkbook.worksheet.table.keys():
        if len(newWorkbook.worksheet.table.get(oneRow, None)) > 0:
            if 'location' == sharedStrings[newWorkbook.worksheet.table[oneRow][0].string]:
                titleRow = oneRow
                for oneCol in newWorkbook.worksheet.table[titleRow].keys():
                    titleDict[sharedStrings[newWorkbook.worksheet.table[titleRow][oneCol].string]] = oneCol
                break
    myChart = newWorkbook.workbook.add_chart({'type': 'column'})
    myChart.add_series({
        'categories': ['Port View', titleRow, titleDict['port'], list(newWorkbook.worksheet.table.keys())[-1], titleDict['port']],
        'values': ['Port View', titleRow, titleDict['lunCount'], list(newWorkbook.worksheet.table.keys())[-1], titleDict['lunCount']],
        'name': 'port'
    })
    newWorkbook.worksheet.insert_chart('B2', myChart)
    '''
    pass


def CreateHostGroupView(arrayJson, newWorkbook):
    '''
    Select the LUNs from each host group
    Extract the required data
    output to Excel tab Host Group View
    '''
    dataFormat = newWorkbook._value1_vertical
    newWorkbook.currentCol = 0
    newWorkbook.currentRow = 4
    newWorkbook.worksheet = newWorkbook.workbook.get_worksheet_by_name("Host Group View")

    newHostGroup = hostGroup()
    classKeys = newHostGroup.__dict__
    titleList = []
    for onKey in classKeys: titleList.append(onKey)
    newWorkbook.addListToRow(titleList, dataFormat) 
    dataFormat = newWorkbook._value1   
    
    for myArray in arrayJson:
            myModel    = myArray['_identity']['model']
            mySerial   = myArray['_identity']['serial']
            myName     = SerialNumber2Name[mySerial]
            myLocation = SerialNumber2Site[mySerial]

            for myPort in myArray['_ports']:
                for myGID in myArray['_ports'][myPort]['_GIDS']:
                    newHostGroup = hostGroup()
                    newHostGroup.serialnumber     = mySerial
                    newHostGroup.arrayModel       = myModel
                    newHostGroup.location         = myLocation
                    newHostGroup.arrayNickname    = myName
                    newHostGroup.port             = myPort
                    newHostGroup.gid              = myGID
                    newHostGroup.loggedin         = ','.join(myArray['_ports'][myPort]['_GIDS'][myGID]['LOGGED_IN'])
                    newHostGroup.hostname         = myArray['_ports'][myPort]['_GIDS'][myGID]['GROUP_NAME']
                    newHostGroup.hostmode         = myArray['_ports'][myPort]['_GIDS'][myGID]['HMD']

                    if '_WWNS' in myArray['_ports'][myPort]['_GIDS'][myGID].keys():
                        myWwns = ','.join(myArray['_ports'][myPort]['_GIDS'][myGID]['_WWNS'].keys())
                        newHostGroup.wwn = myWwns
                    else:
                        newHostGroup.wwn = ''
                        
                    if '_LUNS' in myArray['_ports'][myPort]['_GIDS'][myGID].keys():
                        myLunCount = 0
                        myLunsAlloc = 0
                        myLunsUsed  = 0
                        for myLun in myArray['_ports'][myPort]['_GIDS'][myGID]['_LUNS']: 
                            myLdevNumberDec    = myArray['_ports'][myPort]['_GIDS'][myGID]['_LUNS'][myLun]['LDEV']
                            myLdev = myArray['_ldevlist']['mapped'][myLdevNumberDec]
                            myLunCount +=  1
                            myLunsAlloc += int(float(myLdev['VOL_Capacity(GB)']))
                            myLunsUsed  += int(float(myLdev['Used_Block(GB)']))
                            newHostGroup.rsgid.append(myLdev['RSGID'])
                            newHostGroup.vsmName.append(myArray['_resource_groups'][myLdev['RSGID']]['RS_GROUP'])
                            newHostGroup.virtualSerial.append(myArray['_resource_groups'][myLdev['RSGID']]['V_Serial#'])
                        newHostGroup.lunCount         = myLunCount
                        newHostGroup.LdevCapacityAlloc_GB = myLunsAlloc
                        newHostGroup.LdevCapacityUsed_GB  = myLunsUsed
                        if not newHostGroup.LdevCapacityAlloc_GB == 0:
                            newHostGroup.LdevCapacityUsed_PC = round(((newHostGroup.LdevCapacityUsed_GB / newHostGroup.LdevCapacityAlloc_GB)*100), 2)
                        newHostGroup.rsgid            = ','.join(list(set(newHostGroup.rsgid)))
                        newHostGroup.vsmName          = ','.join(list(set(newHostGroup.vsmName)))
                        newHostGroup.virtualSerial    = ','.join(list(set(newHostGroup.virtualSerial)))
                    else:
                        newHostGroup.rsgid         = ''
                        newHostGroup.vsmName       = ''
                        newHostGroup.virtualSerial = ''

                    # write out the class object to Excel
                    classKeys = newHostGroup.__dict__
                    classList = []
                    for onKey in classKeys: classList.append(classKeys[onKey])
                    newWorkbook.addListToRow(classList, dataFormat)
                             

def CreatePortView(arrayJson, newWorkbook):
    '''
    Select the LUNs from each port
    Extract the required data
    Output to Excel tab 'Port View'
    '''
    global PortChart
    dataFormat = newWorkbook._value1_vertical
    newWorkbook.currentCol = 0
    newWorkbook.currentRow = 4
    newWorkbook.worksheet = newWorkbook.workbook.get_worksheet_by_name("Port View")

    newPort = port()
    classKeys = newPort.__dict__
    titleList = []
    for onKey in classKeys: titleList.append(onKey)
    newWorkbook.addListToRow(titleList, dataFormat)
    dataFormat = newWorkbook._value1

    PortChart = newWorkbook.workbook.add_chart({'type': 'column'})

    mySerialNumbers = [x['_identity']['serial'] for x in arrayJson]
    for myArray in arrayJson:
        myModel    = myArray['_identity']['model']
        mySerial   = myArray['_identity']['serial']
        myName     = SerialNumber2Name[mySerial]
        myLocation = SerialNumber2Site[mySerial]
        newPort = port()
        newPort.arrayModel    = myModel
        newPort.serialnumber  = mySerial
        newPort.arrayNickname = myName
        newPort.location      = myLocation

        myPorts = [x for x in myArray['_ports']]
        for myPort in myPorts:

            newPort = port()
            newPort.arrayModel    = myModel
            newPort.serialnumber  = mySerial
            newPort.arrayNickname = myName
            newPort.location      = myLocation
            newPort.port          = myPort

            myHostGroups = [x for x in myArray['_ports'][myPort]['_GIDS']]
            #newPort.host_groups = str(len(myHostGroups))
            newPort.host_groups = (len(myHostGroups))

            myLuns  = [y for x in myArray['_ports'][myPort]['_GIDS']
                            if '_LUNS' in myArray['_ports'][myPort]['_GIDS'][x]
                            for y in myArray['_ports'][myPort]['_GIDS'][x]['_LUNS']
                        ]
            myLdevs = list(set([myArray['_ports'][myPort]['_GIDS'][x]['_LUNS'][y]['LDEV'] 
                            for x in myArray['_ports'][myPort]['_GIDS']
                            if '_LUNS' in myArray['_ports'][myPort]['_GIDS'][x]
                            for y in myArray['_ports'][myPort]['_GIDS'][x]['_LUNS']
                        ]))
            #newPort.lunCount  = str(len(myLuns))
            #newPort.ldevCount = str(len(myLdevs))
            newPort.lunCount  = (len(myLuns))
            newPort.ldevCount = (len(myLdevs))

            newPort.LdevCapacityAlloc_GB = sum([int(float(myArray['_ldevlist']['mapped'][x]['VOL_Capacity(GB)']))
                        for x in myLdevs
            ])
            newPort.LdevCapacityUsed_GB = sum([int(float(myArray['_ldevlist']['mapped'][x]['Used_Block(GB)']))
                        for x in myLdevs
            ])
            if not newPort.LdevCapacityAlloc_GB == 0:
                newPort.LdevCapacityUsed_PC = round(((newPort.LdevCapacityUsed_GB / newPort.LdevCapacityAlloc_GB)*100), 2)
            
            # write out the class object to Excel
            classKeys = newPort.__dict__
            portList = []
            for onKey in classKeys: portList.append(classKeys[onKey])

            variableRow = newWorkbook.currentRow
            variableName = classKeys['serialnumber'] + "-" + classKeys['port']
            variablePosition = list(classKeys).index('port')
            variableLuns     = list(classKeys).index('lunCount')
            variableSheet    = newWorkbook.worksheet.name

            dataFormat = newWorkbook._value1_CapTB
            newWorkbook.addListToRow(portList, dataFormat)

            PortChart.add_series({
                "name":       f"={variableName}",
                "values":     f"=['{variableSheet}', {variableRow}, {variableLuns}, {variableRow}, {variableLuns}]"
            })
#                "categories": f"=['{variableSheet}', {variableRow}, {variablePosition}, {variableRow}, {variablePosition}]",
            newPort = port()
    PortChart.set_title({"name": "Port Lun Density"})
    PortChart.set_x_axis({"name": "Port"})
    PortChart.set_y_axis({"name": "Count"})
    newWorkbook.worksheet.insert_chart("A1", PortChart)
    return


def CreateLdevLunView(arrayJson, newWorkbook):
    '''
    Select the LUNs from each port
    Extract the required data
    Output to Excel tab 'LDEV LUN View'
    '''
    dataFormat = newWorkbook._value1_vertical
    newWorkbook.currentCol = 0
    newWorkbook.currentRow = 4
    newWorkbook.worksheet = newWorkbook.workbook.get_worksheet_by_name("LDEV LUN View")

    newLedvLun = LdevLun() 
    # print the class titles
    classKeys = newLedvLun.__dict__
    titleList = []
    for onKey in classKeys: titleList.append(onKey)
    newWorkbook.addListToRow(titleList, dataFormat)
    dataFormat = newWorkbook._value1

    for myArray in arrayJson:
        myModel    = myArray['_identity']['model']
        mySerial   = myArray['_identity']['serial']
        myName     = SerialNumber2Name[mySerial]
        myLocation = SerialNumber2Site[mySerial]

        for myPort in myArray['_ports']:
            for myGID in myArray['_ports'][myPort]['_GIDS']:
                if '_WWNS' in myArray['_ports'][myPort]['_GIDS'][myGID].keys():
                    myWwns = ','.join(myArray['_ports'][myPort]['_GIDS'][myGID]['_WWNS'].keys())
                if '_LUNS' in myArray['_ports'][myPort]['_GIDS'][myGID].keys():
                    for myLun in myArray['_ports'][myPort]['_GIDS'][myGID]['_LUNS']:
                        newLedvLun = LdevLun() 
                        newLedvLun.serialnumber     = mySerial
                        newLedvLun.arrayModel       = myModel
                        newLedvLun.location         = myLocation
                        newLedvLun.arrayNickname    = myName
                        newLedvLun.wwn              = myWwns
                        newLedvLun.port             = myPort
                        newLedvLun.gid              = myGID
                        newLedvLun.loggedin         = ','.join(myArray['_ports'][myPort]['_GIDS'][myGID]['LOGGED_IN'])
                        newLedvLun.hostname         = myArray['_ports'][myPort]['_GIDS'][myGID]['GROUP_NAME']
                        newLedvLun.hostmode         = myArray['_ports'][myPort]['_GIDS'][myGID]['HMD']
                        newLedvLun.lunNumber        = myLun
                        newLedvLun.ldevNumberDec    = myArray['_ports'][myPort]['_GIDS'][myGID]['_LUNS'][myLun]['LDEV']
                        newLedvLun.ldevNumberHex    = '{0:x}'.format(int(newLedvLun.ldevNumberDec)).zfill(4)
                        
                        myLdev = myArray['_ldevlist']['mapped'][newLedvLun.ldevNumberDec]
                        
                        newLedvLun.ldevLabel        = myLdev['LDEV_NAMING']
                        newLedvLun.pool             = myLdev['B_POOLID']
                        newLedvLun.LdevCapacityAlloc_GB = float(myLdev['VOL_Capacity(GB)'])
                        newLedvLun.LdevCapacityUsed_GB  = float(myLdev['Used_Block(GB)'])

                        if not newLedvLun.LdevCapacityAlloc_GB == 0:
                            newLedvLun.LdevCapacityUsed_PC = round(((newLedvLun.LdevCapacityUsed_GB / newLedvLun.LdevCapacityAlloc_GB)*100), 2)

                        newLedvLun.volAttr          = ','.join(myLdev['VOL_ATTR']) 
                        newLedvLun.rsgid            = myLdev['RSGID']
                        newLedvLun.vsmName          = myArray['_resource_groups'][newLedvLun.rsgid]['RS_GROUP']
                        newLedvLun.virtualSerial    = myArray['_resource_groups'][newLedvLun.rsgid]['V_Serial#']
                        if 'VIR_LDEV' in myLdev.keys():
                            newLedvLun.virtualLdevNumberDec = myLdev['VIR_LDEV']
                            newLedvLun.virtualLdevNumberHex = '{0:x}'.format(int(newLedvLun.virtualLdevNumberDec)).zfill(4)
                        else:
                            newLedvLun.virtualLdevNumberDec = '-'
                            newLedvLun.virtualLdevNumberHex = '-'
                        # write out the class object to Excel
                        classKeys = newLedvLun.__dict__
                        classList = []
                        for onKey in classKeys: classList.append(classKeys[onKey])
                        newWorkbook.addListToRow(classList, dataFormat)
                        newLedvLun = LdevLun()
    pass


def CreateOsView(arrayJson, newWorkbook):
    '''
    Select the Host Groups from each port
    Extract the required data
    Output to Excel tab 'OS View'
    '''
    dataFormat = newWorkbook._value1_vertical
    newWorkbook.currentCol = 0
    newWorkbook.currentRow = 4
    newWorkbook.worksheet = newWorkbook.workbook.get_worksheet_by_name("OS View")

    newHostMode = hostMode() 
    # print the class titles
    classKeys = newHostMode.__dict__
    titleList = []
    for onKey in classKeys: titleList.append(onKey)
    newWorkbook.addListToRow(titleList, dataFormat)
    dataFormat = newWorkbook._value1

    for myArray in arrayJson:
        myModel    = myArray['_identity']['model']
        mySerial   = myArray['_identity']['serial']
        myName     = SerialNumber2Name[mySerial]
        myLocation = SerialNumber2Site[mySerial]
        myHostModes = sorted(list(set([myArray['_ports'][x]['_GIDS'][y]['HMD'] 
                         for x in myArray['_ports']
                         for y in myArray['_ports'][x]['_GIDS']
                         if '_LUNS' in myArray['_ports'][x]['_GIDS'][y]
        ])))

        # locate the metrics for each host mode
        for oneHostMode in myHostModes:
            # create a new object for each host mode
            newHostMode = hostMode()
            newHostMode.location      = myLocation
            newHostMode.serialnumber  = mySerial
            newHostMode.arrayModel    = myModel
            newHostMode.arrayNickname = myName
            newHostMode.OsName = oneHostMode
            # get the list of host groups matching the host mode
            newHostMode.OsCount = len([myArray['_ports'][x]['_GIDS'][y] 
                         for x in myArray['_ports']
                         for y in myArray['_ports'][x]['_GIDS']
                         if '_LUNS' in myArray['_ports'][x]['_GIDS'][y]
                         if oneHostMode == myArray['_ports'][x]['_GIDS'][y]['HMD']
            ])
            # get the list of ldev ids for the luns that match the host mode
            OsLuns = [myArray['_ports'][x]['_GIDS'][y]['_LUNS'][z]['LDEV']
                         for x in myArray['_ports']
                         for y in myArray['_ports'][x]['_GIDS']
                         if '_LUNS' in myArray['_ports'][x]['_GIDS'][y]
                         if oneHostMode == myArray['_ports'][x]['_GIDS'][y]['HMD']
                         for z in myArray['_ports'][x]['_GIDS'][y]['_LUNS']
            ]
            newHostMode.lunCount = len(OsLuns)
            OsLdevs = sorted(list(set(OsLuns)))
            newHostMode.ldevCount = len(OsLdevs)

            # get the capacity information for each ldev
            for oneLdev in OsLdevs:
                myLdev = myArray['_ldevlist']['mapped'][oneLdev]
                newHostMode.LdevCapacityAlloc_GB += round(float(myLdev['VOL_Capacity(GB)']), 2)
                newHostMode.LdevcapacityUsed_GB  += round(float(myLdev['Used_Block(GB)']),2 )
            
            newHostMode.LdevCapacityAlloc_GB = round(newHostMode.LdevCapacityAlloc_GB, 2)
            newHostMode.LdevcapacityUsed_GB  = round(newHostMode.LdevcapacityUsed_GB, 2)
            if not newHostMode.LdevCapacityAlloc_GB == 0:
                newHostMode.LdevcapacityUsed_PC = round(((newHostMode.LdevcapacityUsed_GB / newHostMode.LdevCapacityAlloc_GB)*100), 2)

            # create output entry for each OS type for a single array
            classKeys = newHostMode.__dict__
            HmdList = []
            for onKey in classKeys: HmdList.append(classKeys[onKey])
            newWorkbook.addListToRow(HmdList, dataFormat)

    return


def CreatePoolView(arrayJson, newWorkbook):
    '''
    Select the pools from each array
    Extract the required data
    Output to Excel tab 'Pool View'
    '''
    dataFormat = newWorkbook._value1_vertical
    newWorkbook.currentCol = 0
    newWorkbook.currentRow = 4
    newWorkbook.worksheet = newWorkbook.workbook.get_worksheet_by_name("Pool View")

    newPool = pool()
    classKeys = newPool.__dict__
    titleList = []
    for onKey in classKeys: titleList.append(onKey)
    newWorkbook.addListToRow(titleList, dataFormat)
    dataFormat = newWorkbook._value1

    for myArray in arrayJson:
        myModel    = myArray['_identity']['model']
        mySerial   = myArray['_identity']['serial']
        myName     = SerialNumber2Name[mySerial]
        myLocation = SerialNumber2Site[mySerial]
        for myIndex in myArray['_pools']:
            myPool = myArray['_pools'][myIndex]
            newPool = pool()
            newPool.arrayModel     = myModel
            newPool.serialnumber   = mySerial
            newPool.arrayNickname  = myName
            newPool.location       = myLocation

            newPool.poolNumber      = myPool['PID']
            newPool.poolName        = myPool['POOL_NAME']
            newPool.poolType        = myPool['PT']
            newPool.poolCapacity_TB = float(myPool['ACT_TP(MB)']) / 1024**2
            newPool.poolFree_TB     = float(myPool['Available(MB)']) / 1024**2
            newPool.poolUsed_TB     = newPool.poolCapacity_TB - newPool.poolFree_TB
            newPool.poolUsed_PC     = myPool['U(%)']
            newPool.poolLdevs       = myPool['LCNT']
            if 'TOTAL_EFF_R' in myPool.keys():
                newPool.poolTotalEfficiecy = myPool['TOTAL_EFF_R']
                newPool.poolDataReduction  = myPool['TLS_R']
                newPool.poolSoftwareSaving = myPool['PLS_R']
                newPool.poolCompression    = myPool['PLS_CMP_R']
                newPool.poolDeduplication  = myPool['PLS_DDP_R']
                newPool.poolPatternMatch   = myPool['PLS_RECLAIM_R']
                newPool.poolProvision_PC   = myPool['PROVISIONING_EFF(%)']
            newPool.poolThreshWarn     = myPool['W(%)']
            newPool.poolThreshDeplete  = myPool['H(%)']

            newPool.poolMapped_TB = sum([float(myArray['_ldevlist']['mapped'][x]['VOL_Capacity(TB)'])
            for x in myArray['_ldevlist']['mapped'] 
            if myArray['_ldevlist']['mapped'][x]['B_POOLID'] == myIndex])

            newPool.poolUnmapped_TB = sum([float(myArray['_ldevlist']['unmapped'][x]['VOL_Capacity(TB)'])
            for x in myArray['_ldevlist']['unmapped'] 
            if 'B_POOLID' in myArray['_ldevlist']['unmapped'][x].keys()
            if myArray['_ldevlist']['unmapped'][x]['B_POOLID'] == myIndex])

            newPool.poolSubscritpion_TB = round((newPool.poolMapped_TB + newPool.poolUnmapped_TB), 2)
            newPool.poolSubscritpion_PC = round(((newPool.poolMapped_TB + newPool.poolUnmapped_TB) / newPool.poolCapacity_TB) * 100, 2)
        
            # create output entry for each pool for a single array
            classKeys = newPool.__dict__
            PoolList = []
            for onKey in classKeys: PoolList.append(classKeys[onKey])
            newWorkbook.addListToRow(PoolList, dataFormat)

    return


def main():
    scriptname = os.path.basename(__file__)
    ts = datetime.now().strftime('%d-%m-%Y_%H.%M.%S')
    logname = '{}_{}'.format(scriptname,ts)
    log = configlog(scriptname,'logs',logname)
    # Validate the input parameters
    parser = argparse.ArgumentParser(description='Produce StorageArray Reports', epilog="")
    parser.add_argument("-i", "--inputfile", help="Source JSON input file path", required=True)
    parser.add_argument("-o", "--outputfile", help="Report file path", required=True)

    args = parser.parse_args()
    source_file = args.inputfile
    output_file = args.outputfile

    # check the source file exists
    if not os.path.exists(source_file):
        log.info('Unable to locate input file: ' + source_file )
        print('Unable to locate input file: ' + source_file)
        exit(666)

    # prefix the data to the output file name and check it does not exist
    currentDate = datetime.now()
    output_path = Path(output_file)
    output_file = '/' + '/'.join(output_path._cparts[1:-1]) + os.sep + currentDate.strftime("%Y%m%d_%H%M") + '__' + output_path.name
    if os.path.exists(output_file):
        log.info('Output file already exists: ' + output_file )
        print('Output file already exists: ' + output_file)
        exit(666)

    # Read the data collected by DataGather.py
    try:
        with open(source_file, 'r')as fh:
            arrayJson = json.load(fh)
    except Exception as myError:
        print('Unable to load cachefile {}'.format(source_file)) 
        exit(4)

    # Create the output file
    output_path = Path(output_file)
    newWorkbook = xlReport.xlReport('/' + '/'.join(output_path._cparts[1:-1]), output_path.name)
    newWorkbook.customerName = 'LBG'
    newWorkbook.documentName = output_path.name
    # create all the required tabs
    newWorkbook.addTOC()
    newWorkbook.addWorksheet('Data Centre View')
    newWorkbook.addWorksheet('Port View')
    newWorkbook.addWorksheet('Host Group View')
    newWorkbook.addWorksheet('LDEV LUN View')
    newWorkbook.addWorksheet('OS View')
    newWorkbook.addWorksheet('Pool View')

    CreatePortView(arrayJson, newWorkbook)
    CreateHostGroupView(arrayJson, newWorkbook)
    ##CreateLdevLunView(arrayJson, newWorkbook)
    CreateOsView(arrayJson, newWorkbook)
    CreatePoolView(arrayJson, newWorkbook)
    CreateDataCentreView(arrayJson, newWorkbook)

    newWorkbook.updateTOC()
    newWorkbook.closeWorkbook()

if __name__ == "__main__":
    main()
    print('-- Successful --')