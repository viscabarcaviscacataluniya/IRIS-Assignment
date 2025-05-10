#!/usr/bin/env python
# coding: utf-8

# In[1]:


#import preferd libraries as per need
from fastapi import FastAPI, Query, HTTPException
from utils.excel_utils import get_sheet_names, get_row_names, get_row_sum
#start of app-> Fast Api
app = FastAPI(
title = "FastAPI Excel Document Parser",
        description="Processes an Excel file and exposes APIs for table listing, row details, and row sum.",
    version="1.0.0")
#global variable -> helps in change's later
excel_path = "capbudg.xls"
#endpoint -> 1:retyrns all sheet names in excel files
@app.get("/list_tables")
def list_tables():
    try:
        sheets = get_sheet_names(excel_path)
        return{"tables": sheets}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
#endpoint -> 2: return forts column va,ues of selected sheets   
@app.get("/get_table_details")
def get_table_details(table_name: str = Query(..., description="Sheet name to fetch row names from")):
    try:
        row_names = get_row_names(excel_path, table_name)
        return {"table_name": table_name, "row_names": row_names} 
    except ValueError as ve:
        raise HTTPException(status_code=404, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

#endpoint-> 3 : return thew sum of rows        
@app.get("/row_sum")
def row_sum(table_name: str = Query(..., description="Sheet name"),row_name: str = Query(..., description="Row name to sum")):
    try:
        total = get_row_sum(excel_path,table_name,row_name)
        return {"table_name": table_name, "row_name": row_name, "sum": total} 
    except ValueError as ve:
            raise HTTPException(status_code=404, detail=str(ve))
    except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))    
            
        
    
    
    





from utils.excel_utils import get_row_names, get_row_sum

excel_path = "capbudg.xls"
sheet_name = "CapBudgWS" 

# Test get_row_names
try:
    row_names = get_row_names(excel_path, sheet_name)
    print(row_names)
except Exception as e:
    print(f"Error: {e}")

# Test get_row_sum
row_name = "Equity Analysis of a Project"  
try:
    row_sum = get_row_sum(excel_path, sheet_name, row_name)
    print(row_sum)
except Exception as e:
    print(f"Error: {e}")




