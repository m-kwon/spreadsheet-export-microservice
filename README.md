# Spreadsheet Export Microservice

A Node.js microservice for exporting RxReceipts data to Excel spreadsheets with embedded images.

## Features

- **Excel Export**: Generate Excel spreadsheets (.xlsx format)
- **Image Embedding**: Include receipt images directly in the spreadsheet
- **Formatted Data**: Properly formatted currency, dates, and categories
- **Filter Support**: Export only filtered/searched receipts

## Installation

```bash
# Install dependencies
npm install

# Start the service
npm start

# Run in development mode
npm run dev

# Run tests
npm test
```

## Dependencies

- **ExcelJS**: Excel file generation with image support
- **Express**: Web framework
- **Axios**: HTTP client for image downloads
- **CORS**: Cross-origin resource sharing

## API Endpoints

### Health Check
```
GET /health
```
Returns service status and configuration.

### Get Export Formats
```
GET /export/formats
```
Returns supported export formats and column information.

### Export Receipts
```
POST /export/receipts
```

**Request Body:**
```json
{
  "receipts": [
    {
      "id": 1,
      "store_name": "CVS Pharmacy",
      "amount": 25.99,
      "receipt_date": "2024-03-15",
      "category": "Pharmacy",
      "description": "Monthly prescription refill",
      "image_id": "550e8400-e29b-41d4-a716-446655440000"
    }
  ],
  "filters": {
    "search": "",
    "category": "Pharmacy",
    "sortBy": "date",
    "sortOrder": "desc"
  },
  "user_info": {
    "name": "John Doe",
    "email": "john@example.com"
  }
}
```

**Response:**
- Content-Type: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- File download with receipt data and embedded images

### Get Metrics
```
GET /export/metrics
```
Returns performance metrics and service capabilities.

## Excel Output Format

The generated Excel file contains:

1. **Title Row**: "Healthcare Receipt Export - [Date]"
2. **Filter Information**: Applied search/category filters (if any)
3. **Header Row**: Column names with professional styling
4. **Data Rows**: Receipt information with the following columns:
   - Store/Provider Name
   - Amount (formatted as currency)
   - Date (formatted as short date)
   - Category
   - Description
   - Receipt Image (embedded image or "No image")
5. **Summary Row**: Total amount and count statistics

## Configuration

### Environment Variables

```bash
PORT=5003
IMAGE_SERVICE_URL=http://localhost:5001
```

### File Structure

```
export-microservice/
├── index.js           # Main service file
├── package.json       # Dependencies and scripts
├── test.js           # Test suite
├── exports/          # Temporary export files (auto-created)
└── test_exports/     # Test output directory
```

## Usage with RxReceipts

1. **Start the Export Service**:
   ```bash
   npm start
   ```

2. **Ensure Image Service is Running** (port 5001)

3. **Use Export Button**: In the RxReceipts app, click "Export to Excel" on the All Receipts page

4. **Download File**: Browser will automatically download the generated Excel file

## Image Embedding

- Images are downloaded from the image service using the `image_id`
- Supported formats: JPEG, PNG
- Images are resized to fit within cells (150x100 pixels)
- Failed image downloads don't stop the export process
- Row height is adjusted to accommodate images

## Integration Example

Frontend integration (React):

```javascript
const handleExport = async () => {
  const exportPayload = {
    receipts: filteredReceipts,
    filters: currentFilters,
    user_info: { name: user.name, email: user.email }
  };

  const response = await fetch('http://localhost:5003/export/receipts', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(exportPayload)
  });

  const blob = await response.blob();
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'receipts_export.xlsx';
  a.click();
};
```