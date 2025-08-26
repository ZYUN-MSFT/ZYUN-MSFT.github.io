/**
 * Simple Excel Table Creator Module
 * A lightweight module that can be integrated into existing Office.js applications
 * 
 * Usage:
 * 1. Include this script after Office.js is loaded
 * 2. Call SimpleTableCreator.init() to initialize
 * 3. Use the provided functions to create tables
 */

const SimpleTableCreator = (function() {
    'use strict';
    
    // Private variables
    let isInitialized = false;
    let isOfficeReady = false;
    
    /**
     * Initialize the table creator
     */
    function init() {
        if (isInitialized) {
            console.log('SimpleTableCreator already initialized');
            return;
        }
        
        Office.onReady((info) => {
            if (info.host === Office.HostType.Excel) {
                isOfficeReady = true;
                console.log('SimpleTableCreator: Office.js ready for Excel');
            } else {
                console.warn('SimpleTableCreator: Not running in Excel');
            }
            isInitialized = true;
        });
    }
    
    /**
     * Check if the module is ready to use
     */
    function isReady() {
        return isInitialized && isOfficeReady;
    }
    
    /**
     * Create a simple table with employee data
     */
    async function createEmployeeTable() {
        if (!isReady()) {
            throw new Error('SimpleTableCreator not ready. Call init() first and ensure running in Excel.');
        }
        
        return Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            
            const data = [
                ['ID', 'Employee Name', 'Department', 'Salary', 'Hire Date'],
                [1, 'Alice Johnson', 'Engineering', 85000, '2022-01-15'],
                [2, 'Bob Smith', 'Marketing', 65000, '2021-05-20'],
                [3, 'Carol Davis', 'Sales', 70000, '2022-03-10'],
                [4, 'David Wilson', 'Engineering', 90000, '2021-11-05'],
                [5, 'Eve Brown', 'HR', 60000, '2020-09-12']
            ];
            
            const range = worksheet.getRange('A1:E6');
            range.clear();
            range.values = data;
            
            const table = worksheet.tables.add(range, true);
            table.name = 'EmployeeData_' + Date.now();
            table.style = 'TableStyleMedium2';
            
            range.format.autofitColumns();
            
            await context.sync();
            return table.name;
        });
    }
    
    /**
     * Create a sales data table
     */
    async function createSalesTable() {
        if (!isReady()) {
            throw new Error('SimpleTableCreator not ready. Call init() first and ensure running in Excel.');
        }
        
        return Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            
            const data = [
                ['Date', 'Product', 'Sales Rep', 'Amount', 'Region'],
                ['2024-01-15', 'Software License', 'John Doe', 12500, 'North'],
                ['2024-01-16', 'Consulting', 'Jane Smith', 8000, 'South'],
                ['2024-01-17', 'Hardware', 'Mike Johnson', 15000, 'East'],
                ['2024-01-18', 'Support', 'Sarah Wilson', 3500, 'West'],
                ['2024-01-19', 'Training', 'Tom Brown', 6000, 'North']
            ];
            
            const range = worksheet.getRange('A1:E6');
            range.clear();
            range.values = data;
            
            const table = worksheet.tables.add(range, true);
            table.name = 'SalesData_' + Date.now();
            table.style = 'TableStyleLight9';
            
            // Format currency column
            const amountColumn = table.columns.getItem('Amount');
            amountColumn.getDataBodyRange().numberFormat = [['$#,##0.00']];
            
            range.format.autofitColumns();
            
            await context.sync();
            return table.name;
        });
    }
    
    /**
     * Create a custom table with specified parameters
     */
    async function createCustomTable(tableName, headers, rowData) {
        if (!isReady()) {
            throw new Error('SimpleTableCreator not ready. Call init() first and ensure running in Excel.');
        }
        
        if (!tableName || !headers || !Array.isArray(headers) || headers.length === 0) {
            throw new Error('Invalid parameters: tableName and headers array are required');
        }
        
        return Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Prepare data array
            const data = [headers];
            if (rowData && Array.isArray(rowData)) {
                data.push(...rowData);
            } else {
                // Generate sample data if no data provided
                for (let i = 0; i < 3; i++) {
                    const row = headers.map((header, index) => `Sample ${header} ${i + 1}`);
                    data.push(row);
                }
            }
            
            // Calculate range
            const endColumn = String.fromCharCode(64 + headers.length); // A=65, but we start from 64 to get correct letter
            const endRow = data.length;
            const rangeAddress = `A1:${endColumn}${endRow}`;
            
            const range = worksheet.getRange(rangeAddress);
            range.clear();
            range.values = data;
            
            const table = worksheet.tables.add(range, true);
            table.name = tableName.replace(/\s+/g, '_') + '_' + Date.now();
            table.style = 'TableStyleMedium6';
            
            range.format.autofitColumns();
            
            await context.sync();
            return table.name;
        });
    }
    
    /**
     * Format an existing table
     */
    async function formatTable(tableName, styleOptions = {}) {
        if (!isReady()) {
            throw new Error('SimpleTableCreator not ready. Call init() first and ensure running in Excel.');
        }
        
        return Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            const table = worksheet.tables.getItem(tableName);
            
            // Apply style
            const style = styleOptions.style || 'TableStyleMedium4';
            table.style = style;
            
            // Format headers if specified
            if (styleOptions.headerColor || styleOptions.headerFontColor) {
                const headerRange = table.getHeaderRowRange();
                if (styleOptions.headerColor) {
                    headerRange.format.fill.color = styleOptions.headerColor;
                }
                if (styleOptions.headerFontColor) {
                    headerRange.format.font.color = styleOptions.headerFontColor;
                }
                headerRange.format.font.bold = true;
            }
            
            // Auto-fit columns
            table.range.format.autofitColumns();
            
            await context.sync();
            return tableName;
        });
    }
    
    /**
     * Get all tables in the active worksheet
     */
    async function getTableNames() {
        if (!isReady()) {
            throw new Error('SimpleTableCreator not ready. Call init() first and ensure running in Excel.');
        }
        
        return Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            const tables = worksheet.tables;
            tables.load(['items/name']);
            
            await context.sync();
            
            return tables.items.map(table => table.name);
        });
    }
    
    /**
     * Delete a table by name
     */
    async function deleteTable(tableName) {
        if (!isReady()) {
            throw new Error('SimpleTableCreator not ready. Call init() first and ensure running in Excel.');
        }
        
        return Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            const table = worksheet.tables.getItem(tableName);
            table.delete();
            
            await context.sync();
            return true;
        });
    }
    
    // Public API
    return {
        init,
        isReady,
        createEmployeeTable,
        createSalesTable,
        createCustomTable,
        formatTable,
        getTableNames,
        deleteTable
    };
})();

// Auto-initialize if Office is already loaded
if (typeof Office !== 'undefined' && Office.onReady) {
    SimpleTableCreator.init();
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SimpleTableCreator;
}

// Make available globally
window.SimpleTableCreator = SimpleTableCreator;