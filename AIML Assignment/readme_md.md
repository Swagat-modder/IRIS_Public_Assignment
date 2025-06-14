# Excel Data Processing API

A FastAPI application that reads data from Excel sheets and provides REST endpoints to interact with the data. The application can parse Excel files, identify tables within sheets, and perform various operations on the data.

## Features

- **Excel File Processing**: Reads and parses Excel files (.xls, .xlsx)
- **Table Identification**: Automatically identifies distinct tables within Excel sheets
- **REST API Endpoints**: Provides clean REST endpoints for data interaction
- **Data Validation**: Handles various data formats including percentages, currencies, and plain numbers
- **Error Handling**: Comprehensive error handling with meaningful error messages
- **API Documentation**: Auto-generated API documentation with Swagger UI

## Installation

### Prerequisites

- Python 3.8 or higher
- pip (Python package installer)

### Setup

1. **Clone or download the project files**

2. **Create a virtual environment (recommended)**
   ```bash
   python -m venv venv
   
   # On Windows
   venv\Scripts\activate
   
   # On macOS/Linux
   source venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Prepare your Excel file**
   - Place your Excel file at `/Data/capbudg.xls`
   - Or modify the `EXCEL_FILE_PATH` variable in `main.py` to point to your file

## Usage

### Starting the Server

```bash
python main.py
```

Or using uvicorn directly:
```bash
uvicorn main:app --host 0.0.0.0 --port 9090
```

The server will start on `http://localhost:9090`

### API Documentation

Once the server is running, you can access:
- **Swagger UI**: `http://localhost:9090/docs`
- **ReDoc**: `http://localhost:9090/redoc`

## API Endpoints

### 1. Root Endpoint
- **URL**: `GET /`
- **Description**: Returns basic API information and available endpoints
- **Response**:
  ```json
  {
    "message": "Excel Data Processing API",
    "version": "1.0.0",
    "endpoints": ["/list_tables", "/get_table_details", "/row_sum"]
  }
  ```

### 2. Health Check
- **URL**: `GET /health`
- **Description**: Check API health and Excel processor status
- **Response**:
  ```json
  {
    "status": "healthy",
    "excel_processor_initialized": true
  }
  ```

### 3. List Tables
- **URL**: `GET /list_tables`
- **Description**: Returns all table names found in the Excel file
- **Response**:
  ```json
  {
    "tables": ["Initial Investment", "Revenue Projections", "Operating Expenses"]
  }
  ```

### 4. Get Table Details
- **URL**: `GET /get_table_details`
- **Parameters**: 
  - `table_name` (query parameter): Name of the table
- **Description**: Returns row names (first column values) for the specified table
- **Example Request**: `GET /get_table_details?table_name=Initial Investment`
- **Response**:
  ```json
  {
    "table_name": "Initial Investment",
    "row_names": [
      "Initial Investment=",
      "Opportunity cost (if any)=",
      "Lifetime of the investment",
      "Salvage Value at end of project=",
      "Deprec. method(1:St.line;2:DDB)=",
      "Tax Credit (if any )=",
      "Other invest.(non-depreciable)="
    ]
  }
  ```

### 5. Row Sum
- **URL**: `GET /row_sum`
- **Parameters**:
  - `table_name` (query parameter): Name of the table
  - `row_name` (query parameter): Name of the row
- **Description**: Calculates the sum of all numerical values in the specified row
- **Example Request**: `GET /row_sum?table_name=Initial Investment&row_name=Tax Credit (if any )=`
- **Response**:
  ```json
  {
    "table_name": "Initial Investment",
    "row_name": "Tax Credit (if any )=",
    "sum": 10.0
  }
  ```

## Data Processing Logic

### Table Identification
The application uses the following logic to identify tables within Excel sheets:
1. Looks for rows where the first cell contains non-empty text
2. Treats these as table headers/names
3. Groups subsequent rows under each table until the next header is found
4. If no clear table structure is found, treats the entire sheet as one table

### Numerical Processing
The row sum calculation handles various data formats:
- **Plain numbers**: `100`, `25.5`
- **Percentages**: `10%`, `5.25%`
- **Currency**: `$1000`, `$25.50`
- **Formatted numbers**: `1,000`, `2,500.75`

Non-numerical values are ignored during sum calculations.

## Testing with Postman

A Postman collection is provided (`postman_collection.json`) with pre-configured requests for all endpoints. To use it:

1. Import the JSON file into Postman
2. The collection includes requests for:
   - API information and health check
   - Listing tables
   - Getting table details for different tables
   - Calculating row sums for various rows

## Error Handling

The API provides comprehensive error handling:

- **404 Not Found**: When a table or row name doesn't exist
- **500 Internal Server Error**: For file processing errors or server issues
- **422 Validation Error**: For invalid query parameters

All errors include descriptive messages to help with debugging.

## Project Structure

```
project/
├── main.py                 # Main FastAPI application
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── postman_collection.json # Postman collection for testing
└── Data/
    └── capbudg.xls        # Excel file to process
```

## Dependencies

- **FastAPI**: Web framework for building APIs
- **Uvicorn**: ASGI server for running FastAPI
- **Pandas**: Data manipulation and analysis
- **OpenPyXL**: Reading Excel files (.xlsx)
- **xlrd**: Reading Excel files (.xls)
- **NumPy**: Numerical computing support

## Potential Improvements

### 1. Enhanced Table Detection
- **Machine Learning Approach**: Use ML algorithms to better identify table boundaries
- **Header Detection**: Improve detection of table headers vs. data rows
- **Multi-level Headers**: Support for complex table structures with nested headers

### 2. Data Processing Enhancements
- **Data Type Inference**: Automatically detect and convert data types
- **Formula Evaluation**: Support for Excel formulas and calculated cells
- **Date/Time Handling**: Better support for date and time formats
- **Statistical Operations**: Add endpoints for mean, median, standard deviation, etc.

### 3. File Format Support
- **Multiple Formats**: Support CSV, ODS, and other spreadsheet formats
- **Multiple Files**: Process multiple Excel files simultaneously
- **File Upload**: Add endpoint for dynamic file uploads
- **Sheet Selection**: Allow users to specify which sheets to process

### 4. Advanced Features
- **Caching**: Implement caching for frequently accessed data
- **Database Integration**: Store processed data in a database
- **Export Functionality**: Export processed data to different formats
- **Data Visualization**: Generate charts and graphs from the data

### 5. User Interface
- **Web Dashboard**: Create a web interface for easier data exploration
- **Real-time Updates**: WebSocket support for real-time data updates
- **User Authentication**: Add user management and authentication

### 6. Performance Optimizations
- **Async Processing**: Use async operations for large file processing
- **Memory Management**: Optimize memory usage for large Excel files
- **Pagination**: Add pagination for large datasets

## Missed Edge Cases

### 1. File-related Edge Cases
- **Empty Excel Files**: Files with no data or only formatting
- **Corrupted Files**: Handling of corrupted or unreadable Excel files
- **Very Large Files**: Memory issues with extremely large Excel files
- **Password Protected Files**: Excel files with password protection
- **Hidden Sheets/Rows**: Data in hidden sheets or rows

### 2. Data Structure Edge Cases
- **Merged Cells**: Excel files with merged cells
- **Complex Layouts**: Tables with irregular structures
- **Multiple Tables per Sheet**: Overlapping or adjacent tables
- **Empty Tables**: Tables with headers but no data
- **Inconsistent Row Names**: Duplicate or similar row names

### 3. Data Type Edge Cases
- **Mixed Data Types**: Rows with both numerical and text data
- **Special Characters**: Unicode characters, symbols, or special formatting
- **Extremely Large Numbers**: Scientific notation or very large numbers
- **Negative Numbers**: Various negative number formats
- **Boolean Values**: TRUE/FALSE values in Excel

### 4. API Usage Edge Cases
- **Concurrent Requests**: Multiple simultaneous requests
- **Invalid Parameters**: Malformed query parameters
- **Special Characters in Names**: Table/row names with special characters
- **Case Sensitivity**: Handling of case differences in names
- **URL Encoding**: Special characters in query parameters

### 5. System Edge Cases
- **Disk Space**: Running out of disk space during processing
- **Memory Limits**: Exceeding available memory
- **Network Issues**: Connection problems during processing
- **Permission Issues**: File access permission problems

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass
6. Submit a pull request

## License

This project is open source and available under the MIT License.

## Support

For questions or issues, please:
1. Check the API documentation at `/docs`
2. Review the error messages for debugging information
3. Ensure your Excel file is properly formatted
4. Verify all dependencies are installed correctly