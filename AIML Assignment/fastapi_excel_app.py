import os
import pandas as pd
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import JSONResponse
from typing import List, Dict, Any, Optional
import numpy as np
from pathlib import Path
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Excel Data Processing API",
    description="A FastAPI application for reading and processing Excel sheet data",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc"
)

class ExcelProcessor:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.excel_data = {}
        self.load_excel_data()
    
    def load_excel_data(self):
        """Load Excel data and identify tables"""
        try:
            if not os.path.exists(self.file_path):
                raise FileNotFoundError(f"Excel file not found: {self.file_path}")
            
            # Read all sheets from the Excel file
            excel_file = pd.ExcelFile(self.file_path)
            
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None)
                self.excel_data[sheet_name] = df
                logger.info(f"Loaded sheet: {sheet_name} with shape {df.shape}")
            
            # Process and identify tables within sheets
            self.process_tables()
            
        except Exception as e:
            logger.error(f"Error loading Excel file: {str(e)}")
            raise HTTPException(status_code=500, detail=f"Error loading Excel file: {str(e)}")
    
    def process_tables(self):
        """Process Excel sheets to identify distinct tables"""
        self.tables = {}
        
        for sheet_name, df in self.excel_data.items():
            # Look for table headers (non-empty cells that might indicate table starts)
            tables_in_sheet = self.identify_tables_in_sheet(df, sheet_name)
            self.tables.update(tables_in_sheet)
    
    def identify_tables_in_sheet(self, df: pd.DataFrame, sheet_name: str) -> Dict[str, pd.DataFrame]:
        """Identify individual tables within a sheet"""
        tables = {}
        
        # Simple approach: look for rows that start with non-null values
        # and seem to be headers (contain text)
        current_table_name = None
        current_table_start = None
        
        for idx, row in df.iterrows():
            # Check if this row might be a table header
            first_cell = row.iloc[0] if len(row) > 0 else None
            
            if pd.notna(first_cell) and isinstance(first_cell, str) and first_cell.strip():
                # This might be a new table
                if current_table_name and current_table_start is not None:
                    # Save the previous table
                    table_df = df.iloc[current_table_start:idx].copy()
                    tables[current_table_name] = self.clean_table(table_df)
                
                # Start new table
                current_table_name = first_cell.strip()
                current_table_start = idx
        
        # Don't forget the last table
        if current_table_name and current_table_start is not None:
            table_df = df.iloc[current_table_start:].copy()
            tables[current_table_name] = self.clean_table(table_df)
        
        # If no clear table structure found, treat the entire sheet as one table
        if not tables:
            table_name = f"Sheet_{sheet_name}" if sheet_name else "Default_Table"
            tables[table_name] = self.clean_table(df)
        
        return tables
    
    def clean_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean and prepare table data"""
        # Remove completely empty rows and columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # Reset index
        df = df.reset_index(drop=True)
        
        return df
    
    def get_table_names(self) -> List[str]:
        """Get list of all table names"""
        return list(self.tables.keys())
    
    def get_table_row_names(self, table_name: str) -> List[str]:
        """Get row names (first column values) for a specific table"""
        if table_name not in self.tables:
            raise HTTPException(status_code=404, detail=f"Table '{table_name}' not found")
        
        df = self.tables[table_name]
        if df.empty:
            return []
        
        # Get first column values, excluding NaN
        row_names = []
        for idx, row in df.iterrows():
            first_cell = row.iloc[0] if len(row) > 0 else None
            if pd.notna(first_cell):
                row_names.append(str(first_cell))
        
        return row_names
    
    def calculate_row_sum(self, table_name: str, row_name: str) -> float:
        """Calculate sum of numerical values in a specific row"""
        if table_name not in self.tables:
            raise HTTPException(status_code=404, detail=f"Table '{table_name}' not found")
        
        df = self.tables[table_name]
        
        # Find the row with the matching row name
        target_row = None
        for idx, row in df.iterrows():
            first_cell = row.iloc[0] if len(row) > 0 else None
            if pd.notna(first_cell) and str(first_cell) == row_name:
                target_row = row
                break
        
        if target_row is None:
            raise HTTPException(status_code=404, detail=f"Row '{row_name}' not found in table '{table_name}'")
        
        # Calculate sum of numerical values (excluding the first column which is the row name)
        numerical_sum = 0.0
        for value in target_row.iloc[1:]:  # Skip first column (row name)
            if pd.notna(value):
                try:
                    # Handle percentage values and other formats
                    if isinstance(value, str):
                        # Remove percentage signs and other non-numeric characters
                        clean_value = value.replace('%', '').replace('$', '').replace(',', '').strip()
                        if clean_value:
                            numerical_sum += float(clean_value)
                    else:
                        numerical_sum += float(value)
                except (ValueError, TypeError):
                    # Skip non-numerical values
                    continue
        
        return numerical_sum

# Initialize the Excel processor
EXCEL_FILE_PATH = "/Data/capbudg.xls"
excel_processor = None

@app.on_event("startup")
async def startup_event():
    """Initialize the Excel processor on startup"""
    global excel_processor
    try:
        excel_processor = ExcelProcessor(EXCEL_FILE_PATH)
        logger.info("Excel processor initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize Excel processor: {str(e)}")
        # For development, we'll create a mock processor if file doesn't exist
        excel_processor = None

@app.get("/")
async def root():
    """Root endpoint with API information"""
    return {
        "message": "Excel Data Processing API",
        "version": "1.0.0",
        "endpoints": [
            "/list_tables",
            "/get_table_details",
            "/row_sum"
        ]
    }

@app.get("/list_tables")
async def list_tables():
    """List all table names present in the Excel sheet"""
    if excel_processor is None:
        raise HTTPException(status_code=500, detail="Excel processor not initialized")
    
    try:
        tables = excel_processor.get_table_names()
        return {"tables": tables}
    except Exception as e:
        logger.error(f"Error listing tables: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error listing tables: {str(e)}")

@app.get("/get_table_details")
async def get_table_details(table_name: str = Query(..., description="Name of the table")):
    """Get row names (first column values) for a specific table"""
    if excel_processor is None:
        raise HTTPException(status_code=500, detail="Excel processor not initialized")
    
    try:
        row_names = excel_processor.get_table_row_names(table_name)
        return {
            "table_name": table_name,
            "row_names": row_names
        }
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error getting table details: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error getting table details: {str(e)}")

@app.get("/row_sum")
async def row_sum(
    table_name: str = Query(..., description="Name of the table"),
    row_name: str = Query(..., description="Name of the row")
):
    """Calculate sum of numerical values in a specific row"""
    if excel_processor is None:
        raise HTTPException(status_code=500, detail="Excel processor not initialized")
    
    try:
        total_sum = excel_processor.calculate_row_sum(table_name, row_name)
        return {
            "table_name": table_name,
            "row_name": row_name,
            "sum": total_sum
        }
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error calculating row sum: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error calculating row sum: {str(e)}")

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {
        "status": "healthy",
        "excel_processor_initialized": excel_processor is not None
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=9090)