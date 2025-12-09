import io
import pandas as pd
from fastapi import FastAPI, File, UploadFile, HTTPException, Body
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import List


app = FastAPI()

#Enable CORS for all origins (for testing purposes) connect with the frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], #replace with frontend URL in production
    allow_methods=["*"],
    allow_headers=["*"],
)

#TEMP STORAGe
TEMP_STORAGE = {}

# 1 - Upload CSV file
@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    if not file.filename.endswith('.csv'):
        raise HTTPException(status_code=400, detail="Only CSV files are supported.")
    try:
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents))
        
        df = df.fillna('')  # Fill NaN values with empty strings
        TEMP_STORAGE['file.filename'] = df
        
        data_preview = df.head().to_dict(orient='split')
        
        return {
            "filename": file.filename,
            "columns": df.columns.tolist(),
            "total_rows": len(data_preview['index']),
            "preview_data": data_preview['data'][:100] # First 100 rows
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")
    
        
        