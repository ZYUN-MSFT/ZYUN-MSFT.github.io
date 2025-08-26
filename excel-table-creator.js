/**
 * Excel Table Creator using Office.js API
 * This module provides functionality to create and manage Excel tables
 */

// Initialize Office.js when ready
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log('Office.js loaded successfully for Excel');
        initializeApp();
    } else {
        console.log('This add-in is designed for Excel');
    }
});

/**
 * Initialize the application
 */
function initializeApp() {
    console.log('Initializing Excel Table Creator');
    
    // Add event listeners when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', setupUI);
    } else {
        setupUI();
    }
}

/**
 * Set up the user interface
 */
function setupUI() {
    const app = document.getElementById('app');
    if (!app) {
        console.error('App container not found');
        return;
    }

    app.innerHTML = `
        <div class="table-creator-container">
            <h2>Excel Table Creator</h2>
            <div class="controls">
                <button id="createSampleTable" class="btn btn-primary">Create Sample Table</button>
                <button id="createCustomTable" class="btn btn-secondary">Create Custom Table</button>
                <button id="formatTable" class="btn btn-info">Format Selected Table</button>
                <button id="clearWorksheet" class="btn btn-warning">Clear Worksheet</button>
            </div>
            
            <div class="custom-table-form" id="customTableForm" style="display: none;">
                <h3>Custom Table Configuration</h3>
                <div class="form-group">
                    <label for="tableName">Table Name:</label>
                    <input type="text" id="tableName" value="CustomTable1" />
                </div>
                <div class="form-group">
                    <label for="tableHeaders">Headers (comma separated):</label>
                    <input type="text" id="tableHeaders" value="Name, Age, City, Department" />
                </div>
                <div class="form-group">
                    <label for="rowCount">Number of rows:</label>
                    <input type="number" id="rowCount" value="5" min="1" max="100" />
                </div>
                <button id="createCustomTableBtn" class="btn btn-success">Create Table</button>
                <button id="cancelCustomTable" class="btn btn-secondary">Cancel</button>
            </div>
            
            <div class="status" id="status"></div>
        </div>
    `;

    // Add basic styling
    addBasicStyles();
    
    // Attach event listeners
    attachEventListeners();
}

/**
 * Add basic styling to the application
 */
function addBasicStyles() {
    const style = document.createElement('style');
    style.textContent = `
        .table-creator-container {
            padding: 20px;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 600px;
            margin: 0 auto;
        }
        
        .controls {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin: 20px 0;
        }
        
        .btn {
            padding: 10px 16px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            transition: background-color 0.2s;
        }
        
        .btn-primary { background-color: #0078d4; color: white; }
        .btn-primary:hover { background-color: #106ebe; }
        
        .btn-secondary { background-color: #6c757d; color: white; }
        .btn-secondary:hover { background-color: #5a6268; }
        
        .btn-info { background-color: #17a2b8; color: white; }
        .btn-info:hover { background-color: #138496; }
        
        .btn-warning { background-color: #ffc107; color: #212529; }
        .btn-warning:hover { background-color: #e0a800; }
        
        .btn-success { background-color: #28a745; color: white; }
        .btn-success:hover { background-color: #218838; }
        
        .custom-table-form {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 4px;
            margin: 20px 0;
        }
        
        .form-group {
            margin-bottom: 15px;
        }
        
        .form-group label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
        }
        
        .form-group input {
            width: 100%;
            padding: 8px 12px;
            border: 1px solid #ced4da;
            border-radius: 4px;
            font-size: 14px;
        }
        
        .status {
            margin-top: 20px;
            padding: 10px;
            border-radius: 4px;
            font-weight: bold;
        }
        
        .status.success {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        
        .status.error {
            background-color: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        
        .status.info {
            background-color: #d1ecf1;
            color: #0c5460;
            border: 1px solid #bee5eb;
        }
    `;
    document.head.appendChild(style);
}

/**
 * Attach event listeners to UI elements
 */
function attachEventListeners() {
    document.getElementById('createSampleTable').addEventListener('click', createSampleTable);
    document.getElementById('createCustomTable').addEventListener('click', showCustomTableForm);
    document.getElementById('formatTable').addEventListener('click', formatSelectedTable);
    document.getElementById('clearWorksheet').addEventListener('click', clearWorksheet);
    document.getElementById('createCustomTableBtn').addEventListener('click', createCustomTable);
    document.getElementById('cancelCustomTable').addEventListener('click', hideCustomTableForm);
}

/**
 * Show status message
 */
function showStatus(message, type = 'info') {
    const status = document.getElementById('status');
    status.textContent = message;
    status.className = `status ${type}`;
    
    // Auto-hide after 5 seconds
    setTimeout(() => {
        status.textContent = '';
        status.className = 'status';
    }, 5000);
}

/**
 * Create a sample table with predefined data
 */
async function createSampleTable() {
    try {
        showStatus('Creating sample table...', 'info');
        
        await Excel.run(async (context) => {
            // Get the active worksheet
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Sample data for the table
            const sampleData = [
                ['Employee ID', 'Name', 'Department', 'Salary', 'Start Date'],
                [1001, 'John Smith', 'Engineering', 75000, '2020-01-15'],
                [1002, 'Sarah Johnson', 'Marketing', 65000, '2019-06-20'],
                [1003, 'Michael Brown', 'Sales', 70000, '2021-03-10'],
                [1004, 'Emily Davis', 'Engineering', 80000, '2020-11-05'],
                [1005, 'David Wilson', 'HR', 60000, '2018-09-12']
            ];
            
            // Clear any existing content in the range
            const range = worksheet.getRange('A1:E6');
            range.clear();
            
            // Add the data to the worksheet
            range.values = sampleData;
            
            // Create a table from the data
            const table = worksheet.tables.add(range, true);
            table.name = 'EmployeeData';
            
            // Apply table style
            table.style = 'TableStyleMedium2';
            
            // Auto-fit columns
            range.format.autofitColumns();
            
            await context.sync();
            
            showStatus('Sample table created successfully!', 'success');
        });
        
    } catch (error) {
        console.error('Error creating sample table:', error);
        showStatus(`Error: ${error.message}`, 'error');
    }
}

/**
 * Show the custom table form
 */
function showCustomTableForm() {
    document.getElementById('customTableForm').style.display = 'block';
}

/**
 * Hide the custom table form
 */
function hideCustomTableForm() {
    document.getElementById('customTableForm').style.display = 'none';
}

/**
 * Create a custom table based on user input
 */
async function createCustomTable() {
    try {
        const tableName = document.getElementById('tableName').value.trim();
        const headersInput = document.getElementById('tableHeaders').value.trim();
        const rowCount = parseInt(document.getElementById('rowCount').value);
        
        if (!tableName || !headersInput) {
            showStatus('Please provide table name and headers', 'error');
            return;
        }
        
        const headers = headersInput.split(',').map(h => h.trim());
        
        if (headers.length === 0) {
            showStatus('Please provide at least one header', 'error');
            return;
        }
        
        showStatus('Creating custom table...', 'info');
        
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Create table data with headers and sample rows
            const tableData = [headers];
            
            // Add sample data rows
            for (let i = 0; i < rowCount; i++) {
                const row = headers.map((header, index) => {
                    // Generate sample data based on header name
                    if (header.toLowerCase().includes('name')) {
                        return `Sample ${header} ${i + 1}`;
                    } else if (header.toLowerCase().includes('id')) {
                        return 1000 + i;
                    } else if (header.toLowerCase().includes('date')) {
                        const date = new Date();
                        date.setDate(date.getDate() + i);
                        return date.toISOString().split('T')[0];
                    } else if (header.toLowerCase().includes('amount') || header.toLowerCase().includes('price') || header.toLowerCase().includes('salary')) {
                        return (Math.random() * 100000).toFixed(2);
                    } else {
                        return `Data ${i + 1}`;
                    }
                });
                tableData.push(row);
            }
            
            // Determine the range for the table
            const endColumn = String.fromCharCode(65 + headers.length - 1); // A=65
            const endRow = rowCount + 1;
            const rangeAddress = `A1:${endColumn}${endRow}`;
            
            // Clear any existing content in the range
            const range = worksheet.getRange(rangeAddress);
            range.clear();
            
            // Add the data to the worksheet
            range.values = tableData;
            
            // Create a table from the data
            const table = worksheet.tables.add(range, true);
            table.name = tableName;
            
            // Apply table style
            table.style = 'TableStyleMedium6';
            
            // Auto-fit columns
            range.format.autofitColumns();
            
            await context.sync();
            
            hideCustomTableForm();
            showStatus(`Custom table "${tableName}" created successfully!`, 'success');
        });
        
    } catch (error) {
        console.error('Error creating custom table:', error);
        showStatus(`Error: ${error.message}`, 'error');
    }
}

/**
 * Format the currently selected table
 */
async function formatSelectedTable() {
    try {
        showStatus('Formatting selected table...', 'info');
        
        await Excel.run(async (context) => {
            // Get the selected range
            const selectedRange = context.workbook.getSelectedRange();
            selectedRange.load(['address', 'worksheet']);
            
            await context.sync();
            
            const worksheet = selectedRange.worksheet;
            
            // Find tables that intersect with the selected range
            const tables = worksheet.tables;
            tables.load(['items']);
            
            await context.sync();
            
            let tableToFormat = null;
            
            // Check each table to see if it intersects with the selection
            for (let i = 0; i < tables.items.length; i++) {
                const table = tables.items[i];
                table.load(['range']);
                await context.sync();
                
                // For simplicity, we'll format the first table found
                // In a real application, you might want more sophisticated logic
                tableToFormat = table;
                break;
            }
            
            if (tableToFormat) {
                // Apply formatting
                tableToFormat.style = 'TableStyleDark1';
                
                // Format headers
                const headerRange = tableToFormat.getHeaderRowRange();
                headerRange.format.fill.color = '#0078d4';
                headerRange.format.font.color = 'white';
                headerRange.format.font.bold = true;
                
                // Format data range
                const dataRange = tableToFormat.getDataBodyRange();
                dataRange.format.font.size = 11;
                
                // Auto-fit columns
                tableToFormat.range.format.autofitColumns();
                
                await context.sync();
                
                showStatus('Table formatted successfully!', 'success');
            } else {
                showStatus('No table found in the selected range. Please select a cell within a table.', 'error');
            }
        });
        
    } catch (error) {
        console.error('Error formatting table:', error);
        showStatus(`Error: ${error.message}`, 'error');
    }
}

/**
 * Clear the current worksheet
 */
async function clearWorksheet() {
    try {
        const confirmed = confirm('Are you sure you want to clear the worksheet? This action cannot be undone.');
        
        if (!confirmed) {
            return;
        }
        
        showStatus('Clearing worksheet...', 'info');
        
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Delete all tables first
            const tables = worksheet.tables;
            tables.load(['items']);
            await context.sync();
            
            tables.items.forEach(table => {
                table.delete();
            });
            
            // Clear all content
            const usedRange = worksheet.getUsedRange();
            if (usedRange) {
                usedRange.clear();
            }
            
            await context.sync();
            
            showStatus('Worksheet cleared successfully!', 'success');
        });
        
    } catch (error) {
        console.error('Error clearing worksheet:', error);
        showStatus(`Error: ${error.message}`, 'error');
    }
}

// Export functions for potential external use
window.ExcelTableCreator = {
    createSampleTable,
    createCustomTable,
    formatSelectedTable,
    clearWorksheet
};