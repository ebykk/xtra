import pandas as pd
pd.set_option('display.width', 2000)
pd.set_option('display.max_columns', 80)
pd.set_option('max_colwidth', 800)

# ------------ INPUTS -----------------------------------------------------------------------------------------------

# Files to convert
targetFileName = 'examples/'+"RVTools_export_all_2021-04-21_10.48.31"
targetFileType = '.xlsx'

# Template to use for converter
templateFileGeneral = "Templates/RVTools_Template_General.xlsx"
templateFile41 = "Templates/RVTools_Template_General_MiB.xlsx"

# Unique Key and sheets to look at
sheetKey = 'server ID' # This is a renamed column that must occur in each tab for the joins (primary key)
# list_of_sheets = ['vInfo']
list_of_sheets = ['vInfo', 'vDisk']

# -------------------------------------------------------------------------------------------------------------------
# Combining inputs
targetFile = targetFileName + targetFileType

# This will return a dictionary of the {SheetName: {ColumnOriginal:ColumnRename}}
# Used to build the individual dataframes with the renamed columns for concat
try:
    def getVersion(FileName):
        df = pd.read_excel(FileName, sheet_name='vMetaData')
        version = df.iloc[0,0]
        print("The version number is: " + str(version))
        return version
    versionNumber = getVersion(targetFile)
except:
    versionNumber = 4.0

# Need to read from the correct template version
if versionNumber == 4.1:
    templateFile = templateFile41
    # list_of_sheets = ['vInfo', 'vDisk']
    # list_of_sheets = ['vInfo']
    print("Used converter template: 4.1")
    print("Sheets used: "+str(list_of_sheets))
else:
    templateFile = templateFileGeneral
    # list_of_sheets = ['vInfo', 'vDisk']
    # list_of_sheets = ['vInfo']
    print("Used converter template: General")
    print("Sheets used: " + str(list_of_sheets))

# Creates a dictionary of dataframes for joining
def getDictColMap(FileName, SheetList):
    dict_if_dfs = {}
    for sheet in SheetList:
        df = pd.read_excel(FileName, sheet_name=sheet, nrows=1, engine='openpyxl')
        df = df.T.dropna()  # transform / drop empty / turn to dictionary
        df.columns = ['Target']
        list1, list2 = list(df.index), list(df.Target)
        zipped = dict(zip(list1, list2))
        dict_if_dfs[sheet]=zipped
    return dict_if_dfs
mappers = getDictColMap(FileName=templateFile, SheetList=list_of_sheets)

# Makes dataframes from sheets -- will perform custom calculation by sheet name if found
def makeFile(TargetFile, SheetList, Key):
    df_list = []
    for sheet in SheetList:
        print()
        print('>>>> Working on %s ...' % sheet)
        cols = list(mappers[sheet].keys())
        df = pd.read_excel(TargetFile, sheet, engine='openpyxl')

        df = df[cols]
        try:
            df.rename(columns=mappers[sheet], inplace=True)
            df.set_index(Key, inplace=True)
        except:
            print('Error')

        if sheet == "vInfo":
            try:
                print('vInfo sub routine used')
                # Storage - the MPA tools requires the RAW storage
                RawName = "Storage-Total Disk Size (MB) (RAID5)"
                df[RawName] = df['Storage Provisioned MB'] * 1.8
                df["Storage Utilization (%) (RAID5)"] = df['Storage In Use MB'] / df[RawName]
                df["Storage Utilization (%) (RAID5)"] = df["Storage Utilization (%) (RAID5)"].apply(
                    lambda x: 1 if x > 1 else x)

                df['CPU-Number of Processors'] = 1
                df['Physical/Virtual'] = 'virtual'
                df['Environment Type'] = 'manually enter'
                df['In Scope of Portfolio'] = df['Template'].apply(lambda x: False if x == True else 'manually enter')
                df['server-Migration Pattern'] = df['Powerstate'].apply(lambda x: 'Retire' if x == 'poweredOff' else '')

                df['OS Name'] = df['OS Version'].map(
                    lambda x: 'Windows' if 'Windows' in x
                    else "Red Hat" if "Red Hat" in x
                    else "SUSE" if "SUSE" in x
                    else "CentOS" if "CentOS" in x
                    else "Linux" if "Linux" in x
                    else "manually enter"
                )

            except:
                print("...vMemory logic had an error but continued...")

        if sheet == "vDisk":
            try:
                print('** vDISK sub routine used')
                df['Storage-Max Read IOPS'] = df['Disk'].apply(lambda x: 251 if "Hard disk" in x else 501)
                df.drop('Disk', axis=1, inplace=True) # needed as the VM might have multiple disks
                df.reset_index(inplace=True)
                # sort so as to drop certain kinds of devices if there are multiple per server (keeps SSD)
                df.sort_values(by='Storage-Max Read IOPS', ascending=True, inplace=True)
                df.drop_duplicates(subset='server ID', keep="last", inplace=True)
                df.set_index(Key, inplace=True)
                # df.to_excel(targetFileName + "_disk.xlsx", sheet_name='disk')
            except:
                print("...vDisk logic had an error but continued...")

        if sheet == "vMemory":
            try:
                print('vMemory sub routine used')
                df["RAM Peak Utilization"] = df['Max RAM Consumed'] / df['RAM-Total Size (MB)']
                df["RAM Peak Utilization"] = df["RAM Peak Utilization"].apply(lambda x: 1 if x > 1 else x)
                print(df["RAM Peak Utilization"].describe())
            except:
                print('...vMemory logic had an error but continued...')

        # print(df.head())
        print('---- Finished %s ...' % sheet)
        df_list.append(df)
    return df_list
list_of_dataframes = makeFile(TargetFile=targetFile, SheetList=list_of_sheets, Key=sheetKey)

def concatDFs(list):
    DFconcat = pd.concat(list, axis=1)
    return DFconcat
DF_concat = concatDFs(list_of_dataframes)

# Reorders the columns to align more closely with the MPA tool
def orderColumns(df):

    coreCols = ['CPU-Number of Processors',
                'CPU-Cores per Processor',
                'RAM-Total Size (MB)',
                'OS Name',
                'OS Version',
                'Hypervisor',
                'Storage-Total Disk Size (MB) (RAID5)',
                'Storage Utilization (%) (RAID5)',
                'Storage-Max Read IOPS',
                'Physical/Virtual',
                'Environment Type',
                'In Scope of Portfolio',
                'server-Migration Pattern'
                ]

    df_coreCols = pd.DataFrame(columns=coreCols)
    df = df_coreCols.append(df)
    df.index.names = [sheetKey]
    return df
DF_concat = orderColumns(DF_concat)

print()
print('>>>> Final DF Preview...')
print(DF_concat.head())

# Colors cells that need to be checked
def colorGeneral(x):
    color = 'yellow' if x == 'manually enter' else 'white'
    return 'background-color: %s' % color
DF_concat = DF_concat.style.applymap(colorGeneral)

def saveXLSX(df):
    df.to_excel(targetFileName+"_output.xlsx", sheet_name='RVTools')
    print()
    print('Saved excel: '+ targetFileName +'_output')
saveXLSX(DF_concat)