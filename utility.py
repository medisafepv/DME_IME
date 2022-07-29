'''
final_PT program

Author: Sion Kim
Contact: sionkim@umich.edu
Latest Edit: 07/29/2022
'''

# Libraries
import io
import re
import pandas as pd
import ipywidgets as widgets
from ipywidgets import HBox, VBox, FileUpload, Layout

class StopExecution(Exception):
    '''
    Silently halts cell execution
    '''
    def _render_traceback_(self):
        pass
    
def prompt_upload(description):
    '''
    Prompts use to upload a file by ipywidget. Returns instance of FileUpload. 
    - description : string desciption for file upload widget
    '''
    uploader = FileUpload(description=description, layout=Layout(width="250px"), multiple=False)
    display(uploader)

    main_display = widgets.Output()

    def on_upload_change(inputs):
        with main_display:
            main_display.clear_output()
            display(list(inputs['new'].keys())[-1])

    uploader.observe(on_upload_change, names='value')
    return uploader

def clean_list(list_in):
    '''
    Removes trailing and leading whitespace from each string element in list
    Maintains order
    '''
    new_list = list()
    for s in list_in:
        new_list.append(s.strip())
        
    return new_list

def manual_identification(columns, missing, filename):
    '''
    Adapted from final_WM program, utility.py 
        Similar behaviour with minor comment changes 
    '''
        
    print("*" * 40)
    print("{} 파일에서 '{}'을 찾지 못했습니다.".format(filename, missing))
    print("{} 파일 제목:".format(filename))
    for i, col in enumerate(columns):
        print("    ({}) {}".format(i, col))
        
    response = input("위에 제목 중 '{}' 컬럼이 있음 (y) 없음 (n) 종료 (q): ".format(missing))
    while response != "y" and response != "n" and response != "q":
        print("잘못 입력. 다시 선택해주세요.")
        response = input("위에 제목 중 '{}' 컬럼이 있음 (y) 없음 (n) 종료 (q): ".format(missing))
            
    if response == "y":
        choice_idx = input("    '{}' 컬럼을 숫자로 선택하세요: ".format(missing))

        while not choice_idx.isdigit():
            choice_idx = input("    '{}' 컬럼을 숫자로 다시 선택하세요 (예: 2): ".format(missing))

        return columns[int(choice_idx)]
    
    if response == "q":
        raise StopExecution
        
    return ""


def confirmation(actual, test):
    '''
    Copied from final_WM program, utility.py 
    '''
        
    response = input("'{}'이 {} 인지 확인 (y/n): ".format(test, actual))
    while response != "y" and response != "n":
        response = input("'{}'이 {} 인지 다시 확인 (y/n): ".format(test, actual))
        
    if response == "y":
        return test
    return ""

def source_identify(columns):
    '''
    Helper function for identifying columns in source data file
    Returns [<Case Number>, <Meddra PT>] in specified order
    '''
    print("=" * 40)
    print("Source data file")
    print("=" * 40)
    

    case_id = ""
    term = ""
    for col in columns:
        if not case_id:
            # if not identified yet
            lower = col.lower()
            if "case" in lower or "number" in lower or "#" in lower:
                case_id = confirmation(actual="Case Number", test=col)
                
        if not term:
            # if not identified yet
            lower = col.lower()
            if "meddra" in lower or "pt" in lower:
                term = confirmation(actual="MedDRRA PT", test=col)
                         
    if not case_id:
        case_id = manual_identification(columns, "Case Number", "Source data")

    if not term:
        term = manual_identification(columns, "MedDRRA PT", "Source data")
        
    # Exeption handling
    while not term:
        term = manual_identification(columns, "MedDRRA PT", "Source data")
        
    return [case_id, term]

def read_excel_raw(uploader, header=0, usecols=None):
    uploader = uploader.value
    file_name = list(uploader.keys())[0]
    return pd.read_excel(io.BytesIO(uploader[file_name]["content"]), header=header, usecols=usecols)

def process_DME(uploader):
    '''
    Reads DME according to version 2020/06/15
    '''
    df = read_excel_raw(uploader, header=15, usecols=[0])
    df.columns = ["PT"]
    return set(df["PT"])

def process_IME(uploader):
    '''
    Reads IME according to version 8 March 2022	
    '''
    df = read_excel_raw(uploader, header=11, usecols=[1])
    df.columns = ["PT"]
    return set(df["PT"])

def process_source(uploader):
    '''
    Adapted from process_ae function in final_WM program
    
    1.1 Added clean list value functionality
    1.2 Removed rows with empty string
    1.3 Prevented reselection if any NaN values present
    '''
    df = read_excel_raw(uploader)
    
    # Drop empty rows at end, if any
    df = df[df.columns].dropna(how="all")
    
    # remove leading/trailing whitespace in column names
    df.columns = clean_list(df.columns)
    
    cols = source_identify(df.columns)
    
    while df[cols[1]].isna().any():
        # If any are NaN values in significant column, reselect
        print("{}컬럼이 비어있습니다. 다시 선택하세요.".format(cols[1]))
        subset = list(df.columns)
        subset.remove(cols[1])
        cols = source_identify(subset)
    
    try:
        # Signficiant column is LLT term or WHO-ART (str), so no type error
        df[cols[1]] = clean_list(list(df[cols[1]])) 
    except AttributeError:
        print("Source 파일 컬럼이 숫자로 인식.")
        print("종료.")
        raise StopExecution
    
    # Remove rows with empty string, if exists after cleaning
    df = df[df[cols[1]].astype(bool)] 
    
    if "" in cols:
        # If missing a column, there can only be missing case no. due to invariant of source_identify() => cols[1]
        df = df.reset_index().rename(columns={"index" : "Row Number"})
        
        df["Row Number"] = df["Row Number"] + 2 
        # Add header and convert to 1-indexing
        
        new_cols = ["Row Number", cols[1]]
        
        return df[new_cols], new_cols
    else:
        return df[cols], cols
    
    
def find_match(data, data_cols, ime, dme):
    '''
    Identifies whether PT in data, column specified data_cols, is in ime or dme
    
    data : source dataframe
    ime : IME set
    dme : DME set
    
    1.1 : Standardized capitalization
    '''
    DME_bool = list()
    IME_bool = list()
    for PT in data[data_cols[1]]:
        if PT.lower() in set(pd.Series(list(dme)).str.lower()):
            DME_bool.append(True)
        else:
            DME_bool.append(False)

        if PT.lower() in set(pd.Series(list(ime)).str.lower()):
            IME_bool.append(True)
        else:
            IME_bool.append(False)
            
    
    final = pd.concat([data, pd.Series(DME_bool, name="DME 해당"), pd.Series(IME_bool, name="IME 해당")], axis=1)
    
    print("IME/DME 해당하는 표시를 선택하세요")
    print("(1) True/False")
    print("(2) O/X")
    print("(3) Custom")
    option = input("Select option: ")
    
    while not option.isdigit() or option not in ["1", "2", "3"]:
        option = input("Reselect option (1, 2, 3): ")
        
    if option == "1":
        return final
    elif option == "2":
        final[["DME 해당", "IME 해당"]] = final[["DME 해당", "IME 해당"]].replace({False : "X", True : "O"})
        return final
    else:
        yes = input("    해당하는 표시: ")
        no = input("    해당하지 않은 표시: ")
        final[["DME 해당", "IME 해당"]] = final[["DME 해당", "IME 해당"]].replace({False : no, True : yes})
        return final
    
def prompt_input():
    words = ["Source Data (.xlsx)", "DME List (.xlsx)", "IME List (.xlsx)"]

    items = []
    for w in words:
        items.append(prompt_upload(description=w))
        
    return items

def control_process(items):
    source, source_cols = process_source(items[0])
    DME = process_DME(items[1])
    IME = process_IME(items[2])
    return find_match(source, source_cols, IME, DME)