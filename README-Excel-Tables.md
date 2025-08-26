# Excel Table Creator with Office.js

This repository contains JavaScript code that demonstrates how to create and manage Excel tables using the Office.js API. The implementation provides both a full-featured UI application and a lightweight module for easy integration.

## 🚀 Features

- **Create Sample Tables**: Generate professional tables with realistic data
- **Custom Table Builder**: Create tables with your own headers and data
- **Table Formatting**: Apply styles, colors, and professional formatting
- **Auto-fit Columns**: Automatically adjust column widths for optimal display
- **Error Handling**: Comprehensive error handling with user-friendly messages
- **Multiple Integration Options**: Choose between full UI or lightweight module

## 📁 Files Overview

### Core Files

1. **`excel-table-creator.js`** - Full-featured table creator with UI
2. **`excel-table-creator.html`** - Complete HTML page with the table creator
3. **`simple-table-creator.js`** - Lightweight module for integration
4. **`manifest-Excel-test-DA.xml`** - Office Add-in manifest file

### Legacy Files
- **`DevXFunnelDataRecorder/`** - Existing add-in project

## 🛠️ Setup and Installation

### Option 1: Standalone HTML Page

1. Open `excel-table-creator.html` in Excel Online or as an Office Add-in
2. The page will automatically load Office.js and initialize the table creator
3. Use the UI controls to create and manage tables

### Option 2: Integration with Existing Project

1. Include the Office.js library in your HTML:
```html
<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
```

2. Include the simple table creator module:
```html
<script src="simple-table-creator.js"></script>
```

3. Initialize and use:
```javascript
// The module auto-initializes, but you can manually initialize if needed
SimpleTableCreator.init();

// Create tables
SimpleTableCreator.createEmployeeTable()
    .then(tableName => console.log('Created table:', tableName))
    .catch(error => console.error('Error:', error));
```

### Option 3: Office Add-in Deployment

1. Use the provided `manifest-Excel-test-DA.xml` as a template
2. Update the source URLs to point to your hosted files
3. Sideload the add-in in Excel following Microsoft's documentation

## 💻 Usage Examples

### Creating a Sample Employee Table

```javascript
// Using the full-featured creator
await createSampleTable();

// Using the simple module
const tableName = await SimpleTableCreator.createEmployeeTable();
console.log('Created table:', tableName);
```

### Creating a Custom Table

```javascript
// Define your custom table
const headers = ['Product', 'Price', 'Category', 'Stock'];
const data = [
    ['Laptop', 999.99, 'Electronics', 15],
    ['Mouse', 29.99, 'Electronics', 50],
    ['Desk', 299.99, 'Furniture', 8]
];

const tableName = await SimpleTableCreator.createCustomTable('ProductInventory', headers, data);
```

### Formatting Tables

```javascript
// Apply custom formatting
await SimpleTableCreator.formatTable('ProductInventory', {
    style: 'TableStyleMedium2',
    headerColor: '#0078d4',
    headerFontColor: 'white'
});
```

## 📊 Available Table Styles

The Office.js API supports various built-in table styles:

- `TableStyleLight1` through `TableStyleLight21`
- `TableStyleMedium1` through `TableStyleMedium28`
- `TableStyleDark1` through `TableStyleDark11`

## 🔧 API Reference

### SimpleTableCreator Methods

#### `init()`
Initialize the table creator module.

#### `isReady()`
Returns `true` if the module is ready to use.

#### `createEmployeeTable()`
Creates a sample employee table with predefined data.

**Returns:** `Promise<string>` - The name of the created table

#### `createSalesTable()`
Creates a sample sales data table.

**Returns:** `Promise<string>` - The name of the created table

#### `createCustomTable(tableName, headers, rowData?)`
Creates a custom table with specified parameters.

**Parameters:**
- `tableName` (string) - Name for the new table
- `headers` (string[]) - Array of column headers
- `rowData` (any[][]) - Optional array of row data

**Returns:** `Promise<string>` - The name of the created table

#### `formatTable(tableName, styleOptions?)`
Formats an existing table.

**Parameters:**
- `tableName` (string) - Name of the table to format
- `styleOptions` (object) - Optional styling options
  - `style` (string) - Table style name
  - `headerColor` (string) - Header background color
  - `headerFontColor` (string) - Header font color

#### `getTableNames()`
Gets all table names in the active worksheet.

**Returns:** `Promise<string[]>` - Array of table names

#### `deleteTable(tableName)`
Deletes a table by name.

**Parameters:**
- `tableName` (string) - Name of the table to delete

**Returns:** `Promise<boolean>` - Success status

## 🎯 Best Practices

1. **Error Handling**: Always wrap table operations in try-catch blocks
2. **Table Names**: Use unique table names to avoid conflicts
3. **Data Validation**: Validate input data before creating tables
4. **Performance**: Use `Excel.run()` to batch operations for better performance
5. **User Feedback**: Provide status updates for long-running operations

## 🐛 Troubleshooting

### Common Issues

1. **"Office.js not loaded"**
   - Ensure the page is running in Excel or Excel Online
   - Check that Office.js CDN is accessible
   - Verify the script tag is correct

2. **"SimpleTableCreator not ready"**
   - Call `SimpleTableCreator.init()` and wait for initialization
   - Check that you're running in Excel (not Word/PowerPoint)

3. **Table creation fails**
   - Ensure the worksheet has enough space for the table
   - Check that table names are unique
   - Verify data array structure is correct

### Debug Mode

Enable debug logging:

```javascript
// Add to your code for debugging
console.log('Office.js ready:', SimpleTableCreator.isReady());
console.log('Available tables:', await SimpleTableCreator.getTableNames());
```

## 🔒 Security Considerations

- Always validate user input before creating tables
- Sanitize table names and data to prevent injection
- Use HTTPS for production deployments
- Follow Microsoft's Office Add-in security guidelines

## 📖 Additional Resources

- [Office.js API Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Excel JavaScript API Reference](https://docs.microsoft.com/en-us/javascript/api/excel)
- [Office Add-ins Development Guide](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly in Excel
5. Submit a pull request

## 📄 License

This project is provided as-is for educational and development purposes. Please ensure compliance with Microsoft's Office Add-in terms of service for production use.

---

**Happy coding with Office.js! 🎉**