from operator import *
from ExcelDoc import ExcelDoc
from TDConnection import TDConnection
from util import *
import os
import xlrd
import operator
import tkMessageBox
import Tkinter as tk
import subprocess
from Tkinter import Grid


#================================================================================
#this section define the global configuration 

HOME_DIR='c:\\Users\\gaoyangao\\OneDrive - DBS Bank Ltd\gaoyan\\05 devl_python'
LIB_DIR=os.path.join(HOME_DIR,'lib')
OUTPUT_DIR=os.path.join(HOME_DIR,'output')
OIA_DD_DIR=os.path.join(HOME_DIR,'oiadd')
OIA_DD_FILE_NAME=os.path.join(OUTPUT_DIR,'oia_dd.csv')
ADM_DD_DIR=os.path.join(HOME_DIR,'admdd')
CDM_DD_DIR=os.path.join(HOME_DIR,'cdmdd')
CDM_DD_FILE_NAME=os.path.join(CDM_DD_DIR,'cdm_dd.csv')
SIM_DD_DIR=os.path.join(HOME_DIR,'simdd')
CDM_MAPPING_DIR=os.path.join(HOME_DIR,'cdmmapping')
CDM_MAPPING_FILE_NAME=os.path.join(OUTPUT_DIR,'cdm_mapping.csv')
SIM_MAPPING_DIR=os.path.join(HOME_DIR,'simmapping')
ADM_MAPPING_DIR=os.path.join(HOME_DIR,'admmapping')
DATA_SOURCE_ID_FILE_NAME=os.path.join(OUTPUT_DIR,'data_source_id.csv')
CDM_PRODUCTION_FILE_NAME=os.path.join(OUTPUT_DIR,'cdm_production_table.csv')
REFERENCE_FILE_NAME=os.path.join(OUTPUT_DIR,'reference_table.csv')
DDURC_OIA_DD_FILE_NAME=os.path.join(OUTPUT_DIR,"ddurc_oiadd.csv")
DDURC_URC_CODE_FILE_NAME=os.path.join(OUTPUT_DIR,'ddurc_urc.csv')
DDURC_CDM_DD_FILE_NAME=os.path.join(OUTPUT_DIR,'ddurc_cdmdd.csv')
DDURC_ADM_DD_FILE_NAME=os.path.join(OUTPUT_DIR,'ddurc_admdd.csv')

DATA_SOURCE_ID_HEADER=["data_source_id", "dta_src_cd", "loc_cd"]
OIA_DD_HEADER= ["Physical Table Name",
        "Physical Field Name",
        "Field Description",
        "Data Type",
        "Nullable"]
OIA_CDM_MAPPING_HEADER = ["Source Table/File",
			"Source Field Name",
			"Target Table/File",
			"Target Field Name",
			"Target Reference Code Table"]

DDURC_DD_HEADER = ["No","Application Code","As of Date",
			"Business Term",
			"Business Term Defintion",
			"Table Name",
			"Column/Field Name",
			"Data Type",
			"Data Length",
			"Mandatory/Optional (M or O)"]
DDURC_URC_HEADER = ["No",
			"Application Code",
			"As of Date",
			"Last Review Date",
			"User Reference Code REF",
			"URC Classification",
			"URC Value",
			"URC Value Description",
			"URC Value Definition",
			"URC Owner(1bank ID)",
			"URC Owner(Department Name)",
			"Maintained By(Department Name)"
			]
#================================================================================

#================================================================================
def refresh_db_object():
	"""
	This will fetch the data from data source id and all the prodution.
	table from DBC
	"""
	iprint("refreshing database object, please wait...")
	myTDConnection = TDConnection( DSN='prod',DBName='bip_vtdb')
	sql_command = """
		select data_source_id, dta_src_cd, loc_cd 
		from bip_vtdb.v0177_DATA_SOURCE_TYPE  
		order by data_source_id
		"""
	dataSourceIdList= myTDConnection.fetchall(sql_command)
	fileName = DATA_SOURCE_ID_FILE_NAME
	csv_writer(dataSourceIdList, fileName, delimiter = ',')	
	sql_command = """
			Select trim(tablename) as tablename from dbc.tables 
			Where databasename = 'BIP_TDB' 
			and trim(tablename) in 
			(Select CDM_Table_Name from bip_cdb.etl_cdm_table 
			Where Table_Type_Cd not in  ('BMAP')) 
			and trim(tablename) like any ('T%','S%') 
			and Trim(tablename) not like 
			All ('%DATA%SOURCE%TYP%','%DATA%SRC%'); 
			"""
	fileName = CDM_PRODUCTION_FILE_NAME
	data= myTDConnection.fetchall(sql_command)
	csv_writer(data, fileName, delimiter = ',')	

	fileName = REFERENCE_FILE_NAME
	data = myTDConnection.fetchall(sql_command = """
			Select trim(table_id) as table_id, 
			'V'||trim((substring(cdm_table_name,2,31))) as table_name,
			trim(cdm_table_name) as original_table_name,
			trim(Model_type) as model_type
			from bip_cdb.ETL_CDM_TABLE
			Where table_type_cd  = 'Reference' and 
			Model_type = 'CDM';
			""")
	csv_writer(data, fileName, delimiter = ',')	
	iprint("refreshing database object complete")
	return True

def getOIAFileCandidate():
	fileName = DATA_SOURCE_ID_FILE_NAME
	dataSourceIdList = csv_reader(fileName)
	DatasourceTemplate=map(lambda x: 'DD_' + str(x[0])+'_'+x[1]+""".xls""", dataSourceIdList)
	oiaFileList = walk(OIA_DD_DIR,r"""(\ADD_.*?)\.xls[x]?$""")
	iprint("the size of oia copybook [{0:4d}]".format(len(oiaFileList)))
	for i,myList in enumerate(oiaFileList):
		if myList[1] not in DatasourceTemplate:
			del oiaFileList[i]
	return oiaFileList

def refresh_oia_dd():
	iprint("refreshing OIA data dictionary, please wait...")
	netDriveConn()
	copytree("Z:",OIA_DD_DIR)
	oiaFileName = OIA_DD_FILE_NAME
	fullData = []
	oiaFileList = getOIAFileCandidate()
	for myPath,myFileName in oiaFileList:
		try:
			excelDoc = ExcelDoc(fileName = myPath,
					tabName = 'DD Data',
					headerLoc=2,
					mode = 'r') 
			valResult = validate_schema(OIA_DD_HEADER,excelDoc.header)
			head = zip(*valResult) # this will get the location of data 
			if not all(head[1]):
				raise  ExcelHeaderError("the header of File of [{}] is not up the standard, please check".format(myPath))

		except ExcelHeaderError as e:
				eprint(e.args)
				eprint("There is issue with the header of File [{}], it will not be processed".format(myPath))

				continue
		except xlrd.XLRDError as e:
				eprint(e.args)
				eprint("There is issues with File of [{}], it might be get corrupted ,it will not be processed".format(myPath))
				continue
		data = excelDoc.getDataOnLoc(head[1])
		empty = False
		for item in data:
			if not item[0] or not item[1]:
				empty = True
				break
		if empty:
			eprint("table name of physical name in OIA file [{}] is empty, please check, this file will not be processed".format(myPath)) 
			continue
		iprint("OIA file [{}] get processed".format(myPath))
		fullData.extend(remove_unicode(data))
		fullData.extend(data)
	fullData.insert(0,OIA_DD_HEADER)
	csv_writer(fullData, oiaFileName, delimiter = ',')	
	iprint("refreshing OIA data dictionary completed")
	return

def refresh_oia_cdm_mapping():
	iprint("refreshing CDM mapping, please wait...")
	netDriveConn()
	copytree("Y:",CDM_MAPPING_DIR)
	fileName = CDM_MAPPING_FILE_NAME
	fullData = []
	fileList = walk(CDM_MAPPING_DIR,r"""(\AT\d{4}[X]?_.*?)_C\.xls[x]?$""")
	iprint("the size of oia to cdm mapping [{0:4d}]".format(len(fileList)))
	fullDict = {}
	ignoreRec = ['NA','N/A','na','n/a']
	for myPath,myFileName in fileList:
		try:
			excelDoc = ExcelDoc(fileName = myPath,
					tabName = 'CDM Mapping',
					headerLoc=1,
					mode = 'r') 
			valResult = validate_schema(OIA_CDM_MAPPING_HEADER,excelDoc.header)
			head = zip(*valResult) # this will get the location of data 
			if not all(head[1]):
				raise  ExcelHeaderError("the header of File {} have issue".format(myPath))

		except ExcelHeaderError as e:
				eprint(e.args)
				eprint("the header of File {} will not be processed".format(myPath))

				continue
		except xlrd.XLRDError as e:
				eprint(e.args)
				eprint("the header of File {} will not be processed".format(myPath))
				continue
		finally:
			pass
		iprint("OIA file {} get processed".format(myPath))
		data = remove_unicode(excelDoc.getDataOnLoc(head[1]))

		for row in data:
			if (not isinstance(row[0],str)) and (not isinstance(row[0],unicode)):
				continue
			if (not isinstance(row[1],str)) and (not isinstance(row[1],unicode)):
				continue
			tableName = row[0].lower()
			columnName = row[1].lower()
			key = tableName + '||' + columnName
			if tableName in ignoreRec:
				continue
			if columnName in ignoreRec:
				continue
			if tableName == "":
				continue
			if columnName == "":
				continue
			fullDict[key] = row

	for k in fullDict:
		fullData.append(fullDict[k])

	fullData.insert(0,OIA_CDM_MAPPING_HEADER)
	csv_writer(fullData, fileName, delimiter = ',')
	iprint("refreshing CDM mapping completed")
	return
#==============================================================================
def create_ddurc_oia_dd():
	iprint("Creating DDURC - OIA Dationary, please wait...")
	oiaDDFileName = OIA_DD_FILE_NAME
	dataSourceIdfileName = DATA_SOURCE_ID_FILE_NAME
	oiaDDData = csv_reader(oiaDDFileName)
	iprint("oiaDD Data file size {:d}".format(len(oiaDDData)))
	dataSourceIdData = csv_reader(dataSourceIdfileName)
	dataSourceIdDataDict = {}
	finalDD = []
	for i in dataSourceIdData[1:]:
		dataSourceIdDataDict[str(i[0])] = i[1]

	for item in oiaDDData: # OIA to CDM mapping is driving
		try:
			businessTerm = item[1];
			businessTermDefintion = item[2]
			tableName = item[0]
			columnName = item[1]
			dataType,dataLength = getDataTypeAndLen(item[3])
			mandatory = 'M' if item[4] == 'N' else 'O'
			sourceSystemCd = dataSourceIdDataDict[getDataSourceCd(item[0])]
		except KeyError as e:
			continue
		finalDD.append([businessTerm,businessTermDefintion,
				tableName,columnName,dataType,
				dataLength,
				mandatory,sourceSystemCd])
	finalDD= sorted(finalDD, key=itemgetter(2,3))
	today = datetime.date.today().strftime("%d-%b-%Y")
	application_cd = "CST-BIP"
	finalDD = [[application_cd,today,f[0],f[1],f[2],f[3],f[4],f[5],f[6]] for f in finalDD]
	map(GenSeq.attachseq,finalDD) 	
	finalDD.insert(0,DDURC_DD_HEADER)
	fileName = DDURC_OIA_DD_FILE_NAME
	csv_writer(finalDD,fileName)
	iprint("Creating DDURC - OIA Dationary compleated")
	return

#========================================================================
def create_ddurc_urc():
	iprint("Creating DDURC - User Reference Code, Please wait...")
	fileName = CDM_DD_FILE_NAME
	data = csv_reader(fileName)
	mapping = {}
	for row in data:
		mapping[row[7]] = row[1]
	
	refTableList = csv_reader(REFERENCE_FILE_NAME)
	prodTableList = csv_reader(CDM_PRODUCTION_FILE_NAME)
	newTotalPivot = zip(*prodTableList)[0]
	desclist = []
	descPattern = r"""_desc"""
	descReg = re.compile(descPattern,re.IGNORECASE)
	conn = TDConnection( DSN='prod',DBName='bip_vtdb')
	for row in refTableList:
		referenceTableName=row[2]
		if referenceTableName in newTotalPivot:
			sqlCmd = "select * from " + "V" +referenceTableName[1:] + ";"
			try:
				fetchRow = conn.fetchall(sql_command = sqlCmd)
				if len(fetchRow[0] ) > 2:
					iprint("table of [{}] have columns other than code and desc, it will not inlcuded as not considered as reference data".format(row[1]))
					continue
				if descReg.search(fetchRow[0][0]):
					map(itemgetter(1,0),fetchRow)
				#fetchRow = [ [mapping[referenceTableName],f[0],f[1] ]for f in fetchRow]

				fetchRow1 = []
				for f in fetchRow[1:]:
					if f[1] == None or f[1] == 'None':
						f[1] = f[0]
					fetchRow1.append([mapping[referenceTableName].upper(),f[0],f[1] ])
				desclist.extend(fetchRow1)
			except: 
				iprint("table of {} not exist".format(row[1]))
				continue
	desclist = remove_unicode(desclist)
	application_cd = "CST-BIP"
	today = datetime.date.today().strftime("%d-%b-%Y")
	urc_classification = "Application"

	desclist = [[application_cd,today,today,f[0],urc_classification,
		f[1],f[2],"","","","",""] for f in desclist]

	desclist = sorted(desclist, key=itemgetter(3,5))
	map(GenSeq.attachseq,desclist) 	
	desclist.insert(0,DDURC_URC_HEADER)
	csv_writer(desclist, DDURC_URC_CODE_FILE_NAME)
	iprint("Creating DDURC - User Reference Code Complete")
	return
#==============================================================================
def CDM_DD_refresh():
	iprint("Refreshing CDM Data Dictionary, please wait...")
	conn= ErwinConn()
        sql_command = """
        SELECT
    	SA.NAME                     'SA Name',
	PEn.name		'Entity Name',
	PEn.definition		'Entity Description',
	PAt.name		'Attribute Name',
	PAt.definition		'Attribute Description',
	PAt.Attribute_order	'Attribute Order',
	PAt.column_order ' Column order',
    Tran(PEn.Physical_Name)     'Table Name',
    Tran(PAt.Physical_Name)     'Column Name',

    PAt.Logical_Data_Type      'Logical Datatype' ,
    PAt.Physical_Data_Type      'Physical Datatype' ,
	PAt.is_logical_only	'logical only',
	PEn.ENTITY_LOGICAL_PROJECT	'Project Name',
	PEn.ENTITY_LOGICAL_PHYSICALIZED	'Physicalize Flag',
	PEn.ENTITY_LOGICAL_FS_SA	'SA Name',
    PEn.ENTITY_LOGICAL_ENTITYTYPE	'Entity Type',
    TRAN(PAt.Null_Option_Type)  'Column Null Option',
    'PK' = CASE
           WHEN TRAN (PAt.Type)  = 'Primary Key'
             THEN 'Yes'
              ELSE 'No'
           END,
    'FK' = CASE
           WHEN IFNULL (PRR.Physical_Name, 'No') = 'No'
            THEN 'No'
             ELSE 'Yes'
           END


    FROM
    Entity PEn
       LEFT JOIN ER_MODEL_SHAPE EMS
         ON EMS.MODEL_OBJECT_REF = PEn.ID@
       LEFT JOIN ER_DIAGRAM ED
         ON ED.ID@ = EMS.OWNER@
       LEFT JOIN SUBJECT_AREA SA
         ON SA.ID@ = ED.OWNER@
       LEFT OUTER JOIN Attribute PAt
          ON PAt.Owner@ = PEn.Id@
       LEFT OUTER JOIN RELATIONSHIP PRR
         ON PAt.Parent_Relationship_Ref = PRR.Id@ 
         where SA.name =  '.DBS 1101 Major Entity'
    Order By 1, 2,7
        """
	fetchRow = conn.fetchall(sql_command)
	iprint ("Total {} of Rows fetched".format(len(fetchRow)))
	csv_writer(fetchRow,CDM_DD_FILE_NAME)
	iprint("Refreshing CDM Data Dictionary Completed")
	return
#==============================================================================

def create_ddurc_cdm_dd():
	iprint ('Creating DDURC CDM Data Dictionary, Please wait...')
	fileName = CDM_DD_FILE_NAME
	data = csv_reader(fileName)
	data = sorted(data, key=itemgetter(7,6))
	iprint("total of {} record retrieved from CDMDD".format(len(data)))
	prodTableList= csv_reader(CDM_PRODUCTION_FILE_NAME)
	iprint("total number of [{}] production tables retrieved".format(len(prodTableList)))
	dddict ={}
	ddList = []
	descReg = re.compile(r"""Data Management.*?Start.*?\n(.*)\n.*Data Management.*End""",re.DOTALL)

	today = datetime.date.today().strftime("%d-%b-%Y")
	application_cd = "CST-BIP"
	infra = ["ins_proc_id","rec_del_flg",
			"st_dt","upd_proc_id","end_dt",
			"as_of_dt","data_source_id","end_ts","st_ts"]

	for row in data:
		if row[11] =="T":
			continue
		if [row[7]] not in prodTableList:
			continue
		if row[1] + row[3] in dddict.keys():
			continue
		dddict[row[1] + row[3]] = True
		businessTerm = row[3]
		newStr=""
		for char in row[4]:
			if char == '\r':
				continue
			newStr+=char

		newDef = descReg.findall(newStr)
		if newDef:
			businessTermDefintion = newDef[0]
		else:
			businessTermDefintion = row[4] 
		tableName = row[7]
		columnName = row[8]
		if columnName in infra:
			continue
		dataType,dataLength = getDataTypeAndLen(row[10])
		mandatory = 'M' if row[16] == 'NOT NULL' else 'O'
		ddList.append([application_cd,
			today,
			businessTerm,
			businessTermDefintion,
			tableName,
			columnName,
			dataType,
			dataLength,
			mandatory])
	ddList = sorted(ddList, key=itemgetter(4,5))
	map(GenSeq.attachseq,ddList) 	
	ddList.insert(0,DDURC_DD_HEADER)

	fileName = DDURC_CDM_DD_FILE_NAME
	csv_writer(remove_unicode(ddList),fileName)
	iprint ('Creating DDURC CDM Data Dictionary Completed')
	return

class Application(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
	self.grid()
        self.createWidgets()
    def createWidgets(self):
	"""
        self.quit_button = tk.Button()
        self.quit_button["text"] = "QUIT"
        self.quit_button["fg"]   = "red"
        self.quit_button["command"] =  self.quit
	"""
	top=self.winfo_toplevel()
        top.rowconfigure(0, weight=1)
        top.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

	self.CheckVar1 = tk.IntVar()
	self.CheckVar2 = tk.IntVar()
	self.CheckVar3 = tk.IntVar()
	self.CheckVar4 = tk.IntVar()
	self.CheckVar5 = tk.IntVar()
	self.CheckVar6 = tk.IntVar()
	self.CheckVar7 = tk.IntVar()
	self.CheckVar8 = tk.IntVar()
	self.var = tk.StringVar()
	self.var.set("Welcome to DD URC Wizard,please select the module to proceed")
	self.L1 = tk.Label(self,textvariable=self.var, \
			anchor = 'w',relief=tk.RAISED ) 	
	self.C1 = tk.Checkbutton(self,text = "DB Object Refresh", 
			variable = self.CheckVar1, \
			onvalue = 1, offvalue = 0, height=2, \
			width = 20,anchor='w')
	self.C2 = tk.Checkbutton(self,text = "OIA DD", variable = self.CheckVar2, \
			onvalue = 1, offvalue = 0, height=2, \
			width = 20,anchor='w')
	self.C3 = tk.Checkbutton(self,text = "CDM DD", variable = self.CheckVar3, \
			onvalue = 1, offvalue = 0, height=2, \
			width = 20,anchor='w')
	self.C4 = tk.Checkbutton(self,text = "OIA CDM Mapping", \
			variable = self.CheckVar4, \
			onvalue = 1, offvalue = 0, height=2, \
			width = 20,anchor='w')
	self.C5 = tk.Checkbutton(self,text = "DD URC OIA DD", \
			variable = self.CheckVar5, \
			onvalue = 1, offvalue = 0, height=2, \
			width = 20,anchor='w')
	self.C6 = tk.Checkbutton(self,text = "DD URC CDM DD", \
			variable = self.CheckVar6, \
			onvalue = 1, offvalue = 0, height=2, \
			width = 20,anchor='w')
	self.C7 = tk.Checkbutton(self,text = "DD URC Reference Code", \
			variable = self.CheckVar7, \
			onvalue = 1, offvalue = 0, height=2, \
			width = 20,anchor='w')
	self.C8 = tk.Checkbutton(self,text = "ALL", \
			variable = self.CheckVar8, \
			onvalue = 1, offvalue = 0, height=2, \
			width = 20,anchor='w')
	self.b1 = tk.Button(self,text="Run",width = 10,command =self.run,anchor='c')
	self.b2 = tk.Button(self,text="Cancel",width = 10,command = quit)
	#self.L1.grid(row=1, column=2, sticky='nswe', pady=2) 
	self.C1.grid(row=2, column=3, sticky='nswe', pady=2) 
	self.C2.grid(row=3, column=3, sticky='nswe', pady=2) 
	self.C3.grid(row=4, column=3, sticky='nswe', pady=2) 
	self.C4.grid(row=5, column=3, sticky='nswe', pady=2) 
	self.C5.grid(row=6, column=3, sticky='nswe', pady=2) 
	self.C6.grid(row=7, column=3, sticky='nswe', pady=2) 
	self.C7.grid(row=8, column=3, sticky='nswe', pady=2) 
	self.b1.grid(row=9, column=2, sticky='nswe', pady=2) 
	self.b2.grid(row=9, column=4, sticky='nswe', pady=2)

    def run(self):
	    if self.CheckVar1.get():
		refresh_db_object()
	    if self.CheckVar2.get():
		refresh_oia_dd()
	    if self.CheckVar3.get():
		CDM_DD_refresh()
	    if self.CheckVar4.get():
		refresh_oia_cdm_mapping()
	    if self.CheckVar5.get():
		create_ddurc_oia_dd()
	    if self.CheckVar6.get():
 		create_ddurc_cdm_dd()
	    if self.CheckVar7.get():
		create_ddurc_urc()

if __name__ == "__main__":
	
	app = Application()
	app.master.title('Welcome To CST-BIP DD-URC Wizard')
	app.mainloop()
