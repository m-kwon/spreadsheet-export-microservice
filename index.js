const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const axios = require('axios');
const path = require('path');
const fs = require('fs').promises;

const app = express();
const PORT = process.env.PORT || 5003;

// Middleware
app.use(cors());
app.use(express.json());

const IMAGE_SERVICE_URL = process.env.IMAGE_SERVICE_URL || 'http://localhost:5001';
const EXPORT_DIR = path.join(__dirname, 'exports');

async function ensureExportDir() {
  try {
    await fs.mkdir(EXPORT_DIR, { recursive: true });
  } catch (error) {
    console.error('Failed to create exports directory:', error);
  }
}

app.get('/health', (req, res) => {
  res.json({
    service: 'Receipt Export Microservice',
    status: 'healthy',
    version: '1.0.0',
    supported_formats: ['Excel (.xlsx)'],
    max_receipts: 1000,
    timestamp: new Date().toISOString()
  });
});

app.get('/export/formats', (req, res) => {
  res.json({
    supported_formats: [
      {
        type: 'Excel',
        extension: 'xlsx',
        mime_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        description: 'Microsoft Excel spreadsheet with receipt data and images',
        features: ['Multiple columns', 'Image embedding', 'Formatting', 'Filtering']
      }
    ],
    columns: [
      { name: 'Store/Provider', field: 'store_name', description: 'Business or healthcare provider name' },
      { name: 'Amount', field: 'amount', description: 'Expense amount in USD' },
      { name: 'Date', field: 'receipt_date', description: 'Date of the expense' },
      { name: 'Category', field: 'category', description: 'Medical expense category' },
      { name: 'Description', field: 'description', description: 'Optional notes about the expense' },
      { name: 'Receipt Image', field: 'image', description: 'Embedded receipt image (if available)' }
    ],
    max_receipts: 1000,
    estimated_processing_time: '5-30 seconds depending on number of images'
  });
});

async function downloadImage(imageId) {
  try {
    if (!imageId) return null;

    const response = await axios.get(`${IMAGE_SERVICE_URL}/image/${imageId}`, {
      responseType: 'arraybuffer',
      timeout: 10000
    });

    if (response.status === 200) {
      return {
        buffer: Buffer.from(response.data),
        contentType: response.headers['content-type'] || 'image/jpeg'
      };
    }
    return null;
  } catch (error) {
    console.error(`Failed to download image ${imageId}:`, error.message);
    return null;
  }
}

function formatCurrency(amount) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD'
  }).format(amount || 0);
}

function formatDate(dateString) {
  try {
    return new Date(dateString).toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  } catch (error) {
    return dateString || 'Invalid Date';
  }
}

app.post('/export/receipts', async (req, res) => {
  const startTime = Date.now();
  let tempFilePath = null;

  try {
    const { receipts, filters = {}, user_info = {} } = req.body;

    if (!receipts || !Array.isArray(receipts)) {
      return res.status(400).json({
        error: 'Invalid request',
        details: 'receipts array is required'
      });
    }

    if (receipts.length === 0) {
      return res.status(400).json({
        error: 'No receipts to export',
        details: 'The receipts array is empty'
      });
    }

    if (receipts.length > 1000) {
      return res.status(400).json({
        error: 'Too many receipts',
        details: 'Maximum 1000 receipts can be exported at once'
      });
    }

    console.log(`Processing export for ${receipts.length} receipts...`);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Receipts Export');

    workbook.creator = user_info.name || 'RxReceipts User';
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.subject = 'Healthcare Receipt Export';
    workbook.description = 'Exported receipt data from RxReceipts application';

    worksheet.columns = [
      { header: 'Store/Provider', key: 'store_name', width: 25 },
      { header: 'Amount', key: 'amount', width: 12 },
      { header: 'Date', key: 'receipt_date', width: 15 },
      { header: 'Category', key: 'category', width: 20 },
      { header: 'Description', key: 'description', width: 40 },
      { header: 'Receipt Image', key: 'image', width: 20 }
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF3498db' }
    };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
    headerRow.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };

    worksheet.insertRow(1, []);
    worksheet.insertRow(1, []);
    worksheet.insertRow(1, [`Healthcare Receipt Export - ${formatDate(new Date().toISOString())}`]);

    const titleRow = worksheet.getRow(1);
    titleRow.font = { bold: true, size: 16, color: { argb: 'FF2c3e50' } };
    titleRow.alignment = { horizontal: 'center' };
    worksheet.mergeCells('A1:F1');

    if (filters.search || filters.category !== 'all') {
      const filterInfo = [];
      if (filters.search) filterInfo.push(`Search: "${filters.search}"`);
      if (filters.category && filters.category !== 'all') filterInfo.push(`Category: ${filters.category}`);

      worksheet.insertRow(2, [`Filters Applied: ${filterInfo.join(', ')}`]);
      const filterRow = worksheet.getRow(2);
      filterRow.font = { italic: true, color: { argb: 'FF7f8c8d' } };
      worksheet.mergeCells('A2:F2');
    }

    let currentRow = 4;
    let imageCounter = 0;

    for (let i = 0; i < receipts.length; i++) {
      const receipt = receipts[i];

      const row = worksheet.addRow({
        store_name: receipt.store_name || '',
        amount: formatCurrency(receipt.amount),
        receipt_date: formatDate(receipt.receipt_date),
        category: receipt.category || '',
        description: receipt.description || '',
        image: receipt.image_id ? 'Image attached' : 'No image'
      });

      row.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
      row.border = {
        top: { style: 'thin', color: { argb: 'FFe0e0e0' } },
        left: { style: 'thin', color: { argb: 'FFe0e0e0' } },
        bottom: { style: 'thin', color: { argb: 'FFe0e0e0' } },
        right: { style: 'thin', color: { argb: 'FFe0e0e0' } }
      };

      if (receipt.image_id) {
        try {
          console.log(`Downloading image ${i + 1}/${receipts.length}: ${receipt.image_id}`);

          const imageData = await downloadImage(receipt.image_id);
          if (imageData) {
            const imageId = workbook.addImage({
              buffer: imageData.buffer,
              extension: imageData.contentType.includes('png') ? 'png' : 'jpeg'
            });

            row.height = 100;

            worksheet.addImage(imageId, {
              tl: { col: 5, row: currentRow - 1 },
              ext: { width: 150, height: 100 }
            });

            imageCounter++;
          }
        } catch (imageError) {
          console.error(`Failed to process image for receipt ${receipt.id}:`, imageError.message);
        }
      }

      currentRow++;
    }

    const summaryRow = worksheet.addRow([
      'TOTAL:',
      formatCurrency(receipts.reduce((sum, r) => sum + (parseFloat(r.amount) || 0), 0)),
      `${receipts.length} receipts`,
      '',
      '',
      `${imageCounter} images`
    ]);

    summaryRow.font = { bold: true };
    summaryRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFecf0f1' }
    };

    for (let i = 1; i <= 5; i++) {
      worksheet.getColumn(i).width = Math.max(worksheet.getColumn(i).width || 0, 15);
    }

    const timestamp = new Date().toISOString().split('T')[0];
    const filename = `receipts_export_${timestamp}_${Date.now()}.xlsx`;
    tempFilePath = path.join(EXPORT_DIR, filename);

    console.log('Generating Excel file...');
    await workbook.xlsx.writeFile(tempFilePath);

    const processingTime = Date.now() - startTime;
    console.log(`Export completed in ${processingTime}ms with ${imageCounter} images embedded`);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('X-Processing-Time', processingTime.toString());
    res.setHeader('X-Images-Included', imageCounter.toString());

    const fileStream = require('fs').createReadStream(tempFilePath);
    fileStream.pipe(res);

    fileStream.on('end', async () => {
      try {
        await fs.unlink(tempFilePath);
        console.log(`Cleaned up temporary file: ${filename}`);
      } catch (cleanupError) {
        console.error('Failed to clean up temporary file:', cleanupError);
      }
    });

  } catch (error) {
    console.error('Export error:', error);

    if (tempFilePath) {
      try {
        await fs.unlink(tempFilePath);
      } catch (cleanupError) {
      }
    }

    const processingTime = Date.now() - startTime;

    res.status(500).json({
      success: false,
      error: 'Export failed',
      details: error.message,
      processing_time_ms: processingTime,
      timestamp: new Date().toISOString()
    });
  }
});

app.get('/export/metrics', (req, res) => {
  res.json({
    service: 'Receipt Export Microservice',
    status: 'operational',
    supported_formats: ['xlsx'],
    performance: {
      max_receipts: 1000,
      estimated_time_per_receipt: '50-200ms',
      estimated_time_per_image: '500-2000ms',
      typical_file_size: '500KB - 50MB (depending on images)'
    },
    features: [
      'Excel spreadsheet generation',
      'Receipt image embedding',
      'Formatted currency and dates',
      'Filter information inclusion',
      'Summary calculations',
      'Professional styling'
    ],
    timestamp: new Date().toISOString()
  });
});

app.use((error, req, res, next) => {
  console.error('Unhandled error:', error);
  res.status(500).json({
    error: 'Internal server error',
    details: error.message,
    timestamp: new Date().toISOString()
  });
});

app.use((req, res) => {
  res.status(404).json({
    error: 'Endpoint not found',
    available_endpoints: [
      'GET /health',
      'GET /export/formats',
      'POST /export/receipts',
      'GET /export/metrics'
    ]
  });
});

async function startServer() {
  try {
    await ensureExportDir();

    app.listen(PORT, () => {
      console.log(`Receipt Export Microservice running on port ${PORT}`);
      console.log(`Health check: http://localhost:${PORT}/health`);
      console.log(`Export endpoint: http://localhost:${PORT}/export/receipts`);
      console.log('Supported formats: Excel (.xlsx)');
    });
  } catch (error) {
    console.error('Failed to start server:', error);
    process.exit(1);
  }
}

startServer();

module.exports = app;