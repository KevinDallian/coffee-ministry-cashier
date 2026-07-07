const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const moment = require('moment');

// --- Configuration Constants ---
const ORDERS_FILE = './orders.json';
const EXPORT_DIR = './exports';
const EXPORT_FILE_NAME = 'orders-report.xlsx'; // Keep for base name, though mostly constructed in exportExcel

// Canonical list of menu items for consistent reporting
const MENU_ITEMS_CANONICAL = [
    'Hot Latte', 'Hot Americano', 'Hot Cappuccino', 'Hot Espresso',
    'Ice Latte', 'Ice Americano', 'Milk'
];

// --- Helper Functions ---

/**
 * Loads orders data from the orders.json file.
 * Handles file not found or JSON parsing errors.
 * @returns {Array<Object>} An array of order objects. Returns an empty array if an error occurs.
 */
function loadOrders() {
    try {
        const data = fs.readFileSync(ORDERS_FILE, 'utf8');
        return JSON.parse(data);
    } catch (error) {
        if (error.code === 'ENOENT') {
            console.error(`Error: orders.json not found at ${ORDERS_FILE}`);
        } else {
            console.error('Error loading or parsing orders.json:', error.message);
        }
        return [];
    }
}

/**
 * Normalizes a menu item name to its canonical form (case-insensitive,
 * treats "Espresso" and "Hot Espresso" as the same).
 * @param {string} name - The original menu item name.
 * @returns {string} The normalized menu item name, or the original name if it's not a known canonical item.
 */
function normalizeMenuName(name) {
    const lowerName = name.toLowerCase();
    if (lowerName === 'espresso') {
        return 'Hot Espresso'; // Treat "Espresso" as "Hot Espresso"
    }
    for (const canonical of MENU_ITEMS_CANONICAL) {
        if (canonical.toLowerCase() === lowerName) {
            return canonical;
        }
    }
    return name; // Return original if no match (for 'Other Items')
}

/**
 * Calculates dashboard statistics such as total, completed, pending orders,
 * and completion rate.
 * @param {Array<Object>} orders - The array of order objects.
 * @returns {Object} An object containing dashboard statistics.
 */
function calculateDashboard(orders) {
    const totalOrders = orders.length;
    const completedOrders = orders.filter(o => o.status === 'done').length;
    const pendingOrders = totalOrders - completedOrders;
    // Calculate completion rate, handling division by zero
    const completionRate = totalOrders > 0 ? (completedOrders / totalOrders * 100).toFixed(2) : '0.00';
    const exportDate = moment().format('YYYY-MM-DD HH:mm:ss');

    return {
        totalOrders,
        completedOrders,
        pendingOrders,
        completionRate,
        exportDate
    };
}

/**
 * Calculates summary statistics for all canonical menu items, including total quantity
 * and percentage distribution.
 * @param {Array<Object>} orders - The array of order objects.
 * @param {Array<string>} menuItemsCanonical - Canonical list of menu item names.
 * @returns {Object} An object containing an array of summary items and the total quantity of canonical drinks.
 */
function calculateSummary(orders, menuItemsCanonical) {
    const itemQuantities = {};
    let totalDrinkQty = 0;

    // Initialize quantities for all canonical items to zero
    menuItemsCanonical.forEach(item => {
        itemQuantities[item] = 0;
    });

    // Aggregate quantities for canonical items from all orders
    orders.forEach(order => {
        order.items.forEach(item => {
            const normalizedName = normalizeMenuName(item.name);
            if (menuItemsCanonical.includes(normalizedName)) {
                itemQuantities[normalizedName] += item.qty;
                totalDrinkQty += item.qty;
            }
        });
    });

    // Format summary data with quantities and percentages
    const summaryData = menuItemsCanonical.map(item => {
        const quantity = itemQuantities[item];
        const percentage = totalDrinkQty > 0 ? (quantity / totalDrinkQty * 100).toFixed(2) : '0.00';
        return { item, quantity, percentage: `${percentage}%` };
    });

    return { summaryData, totalDrinkQty };
}

/**
 * Identifies and tallies quantities for menu items that are not in the canonical list.
 * These are categorized as 'Other Items'.
 * @param {Array<Object>} orders - The array of order objects.
 * @param {Array<string>} menuItemsCanonical - Canonical list of menu item names.
 * @returns {Array<Object>} An array of other item objects with their name and aggregated quantity.
 */
function calculateOtherItems(orders, menuItemsCanonical) {
    const otherItemQuantities = {};

    orders.forEach(order => {
        order.items.forEach(item => {
            const normalizedName = normalizeMenuName(item.name);
            if (!menuItemsCanonical.includes(normalizedName)) {
                otherItemQuantities[item.name] = (otherItemQuantities[item.name] || 0) + item.qty;
            }
        });
    });

    return Object.entries(otherItemQuantities).map(([item, quantity]) => ({ item, quantity }));
}

/**
 * Builds the Excel worksheet including dashboard, main orders table, summary, and
 * other items sections, applying all specified formatting and printing configurations.
 * @param {ExcelJS.Workbook} workbook - The ExcelJS workbook object.
 * @param {Array<Object>} orders - The sorted array of order objects.
 * @param {Object} dashboardData - Dashboard statistics.
 * @param {Array<Object>} summaryData - Summary item data.
 * @param {Array<Object>} otherItemsData - Other items data.
 * @param {Array<string>} menuItemsCanonical - Canonical list of menu item names.
 */
function buildWorksheet(workbook, orders, dashboardData, summaryData, otherItemsData, menuItemsCanonical) {
    const worksheet = workbook.addWorksheet('Orders Report');

    // --- Workbook Metadata ---
    workbook.creator = 'Cashier System';
    workbook.company = 'Local Coffee Shop';
    workbook.subject = 'Daily Order Report';

    // --- Common Styles ---
    const boldFont = { name: 'Arial', bold: true, size: 10 };
    const titleFont = { name: 'Arial', size: 16, bold: true };
    const thinBorderStyle = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
    };

    // Default font for the entire worksheet (managed by specific cell styles now)
    worksheet.properties.defaultColWidth = 15; // Set a reasonable default width
    worksheet.properties.defaultRowHeight = 20; // Set a reasonable default height

    const headerStyle = {
        font: { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' } }, // White font
        alignment: { vertical: 'middle', horizontal: 'center' },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF264E13' } }, // Dark Green fill
        border: thinBorderStyle
    };
    const textCellStyle = {
        font: { name: 'Arial', size: 10 },
        alignment: { vertical: 'top', horizontal: 'left' },
        border: thinBorderStyle
    };
    const numberCellStyle = {
        font: { name: 'Arial', size: 10 },
        alignment: { vertical: 'top', horizontal: 'center' },
        border: thinBorderStyle
    };
    const dateCellStyle = {
        font: { name: 'Arial', size: 10 },
        alignment: { vertical: 'top', horizontal: 'left' },
        border: thinBorderStyle,
        numFmt: 'hh:mm:ss' // Changed to show only hour:minute:second
    };
    const wrapTextStyle = { alignment: { wrapText: true, vertical: 'top', horizontal: 'left' } };

    // Alternating row colors
    const alternatingFill1 = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E7DF' } }; // Light Grey/Green
    const alternatingFill2 = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } }; // White

    const totalRowFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } }; // Light Grey


    let currentRow = 1;

    // --- Main Table Headers ---
    const mainTableHeaderRow = currentRow;
    const mainTableColumnsDefinition = [
        { header: 'No', key: 'no' },
        { header: 'Name', key: 'name' },
        ...menuItemsCanonical.map(item => ({ header: item, key: item.toLowerCase().replace(/ /g, '') })),
        { header: 'Notes', key: 'notes' },
        { header: 'Status', key: 'status' },
        { header: 'Created At', key: 'createdAt' },
        { header: 'Updated At', key: 'updatedAt' }
    ];

    // Set columns for the worksheet
    worksheet.columns = mainTableColumnsDefinition;

    // Apply header style to the main table header row
    worksheet.getRow(mainTableHeaderRow).eachCell((cell) => {
        cell.style = headerStyle;
    });

    // Apply auto-filter and freeze pane on the header row
    worksheet.autoFilter = {
        from: { row: mainTableHeaderRow, column: 1 },
        to: { row: mainTableHeaderRow, column: mainTableColumnsDefinition.length }
    };
    worksheet.views = [{ state: 'frozen', ySplit: mainTableHeaderRow }];
    currentRow++;

    // --- Main Table Data ---
    orders.forEach((order, index) => {
        const row = worksheet.getRow(currentRow);
        row.getCell('no').value = index + 1;
        row.getCell('name').value = order.tableLabel;
        row.getCell('notes').value = order.note;
        row.getCell('notes').alignment = wrapTextStyle.alignment; // Apply wrap text for Notes column
        row.getCell('status').value = order.status;
        row.getCell('createdAt').value = moment(order.createdAt).toDate(); // Use Date object for Excel formatting
        row.getCell('createdAt').numFmt = dateCellStyle.numFmt;

        if (order.updatedAt) {
            row.getCell('updatedAt').value = moment(order.updatedAt).toDate();
            row.getCell('updatedAt').numFmt = dateCellStyle.numFmt;
        } else {
            row.getCell('updatedAt').value = ''; // Leave empty if updatedAt is missing
        }

        // Aggregate item quantities for the current order
        const orderItemQuantities = {};
        order.items.forEach(item => {
            const normalizedName = normalizeMenuName(item.name);
            orderItemQuantities[normalizedName] = (orderItemQuantities[normalizedName] || 0) + item.qty;
        });

        // Populate menu item columns
        menuItemsCanonical.forEach(menuItem => {
            const key = menuItem.toLowerCase().replace(/ /g, '');
            if (orderItemQuantities[menuItem]) {
                row.getCell(key).value = orderItemQuantities[menuItem];
                row.getCell(key).alignment = numberCellStyle.alignment;
            } else {
                row.getCell(key).value = ''; // Ensure empty cells for non-ordered items
            }
        });

        // Apply general cell style and alternating background fill for the entire row
        const rowFill = index % 2 === 0 ? alternatingFill2 : alternatingFill1; // Alternating white and #e2e7df

        row.eachCell((cell, colNumber) => {
            // Start with base text cell style (includes font, borders, default alignment)
            cell.style = { ...textCellStyle }; // Reapply base style
            cell.font = textCellStyle.font; // Explicitly set font from textCellStyle

            // Apply number cell style for quantity columns if they have a numeric value
            if (colNumber > 2 && colNumber <= (2 + mainTableColumnsDefinition.length) && typeof cell.value === 'number') {
                cell.alignment = numberCellStyle.alignment;
                cell.font = numberCellStyle.font;
            } else if (colNumber === mainTableColumnsDefinition.findIndex(col => col.key === 'notes') + 1) { // Notes column
                cell.alignment = wrapTextStyle.alignment;
            } else if (colNumber === mainTableColumnsDefinition.findIndex(col => col.key === 'createdAt') + 1 ||
                       colNumber === mainTableColumnsDefinition.findIndex(col => col.key === 'updatedAt') + 1) {
                cell.numFmt = dateCellStyle.numFmt; // Ensure date format for date columns
                cell.font = dateCellStyle.font;
            }

            // Apply the alternating row fill last to ensure it takes precedence
            cell.fill = rowFill;
        });
        // Override alignment for 'Notes' column after all other styles are applied
        row.getCell('notes').alignment = wrapTextStyle.alignment;
        currentRow++;
    });
    currentRow = 2;

    // --- Summary Section ---
    const summaryStartCol = 15; // Column O
    worksheet.getCell(currentRow, summaryStartCol).value = 'Summary';
    worksheet.getCell(currentRow, summaryStartCol).font = { name: 'Arial', bold: true, size: 10 };
    currentRow++;

    worksheet.getCell(currentRow, summaryStartCol).value = 'Item';
    worksheet.getCell(currentRow, summaryStartCol + 1).value = 'Quantity';
    worksheet.getCell(currentRow, summaryStartCol + 2).value = 'Percentage';
    worksheet.getRow(currentRow).eachCell((cell, colNumber) => {
        if (colNumber >= summaryStartCol && colNumber <= summaryStartCol + 2) {
            cell.style = headerStyle;
        }
    });
    currentRow++;

    summaryData.forEach((data, index) => {
        const row = worksheet.getRow(currentRow);
        row.getCell(summaryStartCol).value = data.item;
        row.getCell(summaryStartCol + 1).value = data.quantity;
        row.getCell(summaryStartCol + 2).value = data.percentage;
        const rowFill = index % 2 === 0 ? alternatingFill2 : alternatingFill1; // Alternating white and #e2e7df
        row.eachCell((cell, colNumber) => {
            if (colNumber >= summaryStartCol && colNumber <= summaryStartCol + 2) {
                cell.style = { ...textCellStyle }; // Start with base text style
                cell.font = textCellStyle.font; // Explicitly set font
                if (colNumber === summaryStartCol + 1 || colNumber === summaryStartCol + 2) { // Quantity and Percentage columns
                    cell.alignment = numberCellStyle.alignment;
                    cell.font = numberCellStyle.font;
                }
                cell.fill = rowFill; // Apply alternating fill
            }
        });
        currentRow++;
    });

    // Total row for Summary
    const totalSummaryRow = worksheet.getRow(currentRow);
    totalSummaryRow.getCell(summaryStartCol).value = 'Total';
    totalSummaryRow.getCell(summaryStartCol + 1).value = summaryData.reduce((acc, curr) => acc + curr.quantity, 0); // Sum of canonical item quantities
    totalSummaryRow.getCell(summaryStartCol + 2).value = '100.00%';
    totalSummaryRow.eachCell((cell, colNumber) => {
        if (colNumber >= summaryStartCol && colNumber <= summaryStartCol + 2) {
            cell.font = { name: 'Arial', bold: true, size: 10 }; // Apply Arial bold font
            cell.fill = totalRowFill; // Apply total row fill
            cell.style = { ...textCellStyle }; // Apply base text style for borders
            if (colNumber === summaryStartCol + 1 || colNumber === summaryStartCol + 2) {
                cell.alignment = numberCellStyle.alignment;
            }
        }
    });
    currentRow += 2;

    // --- Other Items Section ---
    if (otherItemsData.length > 0) {
        const otherItemsStartCol = 15; // Column O
        worksheet.getCell(currentRow, otherItemsStartCol).value = 'Other Items';
        worksheet.getCell(currentRow, otherItemsStartCol).font = { name: 'Arial', bold: true, size: 10 };
        currentRow++;

        worksheet.getCell(currentRow, otherItemsStartCol).value = 'Other Item';
        worksheet.getCell(currentRow, otherItemsStartCol + 1).value = 'Quantity';
        worksheet.getRow(currentRow).eachCell((cell, colNumber) => {
            if (colNumber >= otherItemsStartCol && colNumber <= otherItemsStartCol + 1) {
                cell.style = headerStyle;
            }
        });
        currentRow++;

        otherItemsData.forEach((data, index) => {
            const row = worksheet.getRow(currentRow);
            row.getCell(otherItemsStartCol).value = data.item;
            row.getCell(otherItemsStartCol + 1).value = data.quantity;
            const rowFill = index % 2 === 0 ? alternatingFill2 : alternatingFill1; // Alternating white and #e2e7df
            row.eachCell((cell, colNumber) => {
                if (colNumber >= otherItemsStartCol && colNumber <= otherItemsStartCol + 1) {
                    cell.style = { ...textCellStyle }; // Start with base text style
                    cell.font = textCellStyle.font; // Explicitly set font
                    if (colNumber === otherItemsStartCol + 1) { // Quantity column
                        cell.alignment = numberCellStyle.alignment;
                        cell.font = numberCellStyle.font;
                    }
                    cell.fill = rowFill; // Apply alternating fill
                }
            });
            currentRow++;
        });
    }

    // --- Printing Configuration ---
    worksheet.pageSetup.orientation = 'landscape';
    worksheet.pageSetup.fitToPage = true;
    worksheet.pageSetup.fitToWidth = 1;
    // Repeat the main table header row on every printed page
    worksheet.pageSetup.printTitlesRow = `${mainTableHeaderRow}:${mainTableHeaderRow}`;
    worksheet.pageSetup.margins = {
        left: 0.7, right: 0.7,
        top: 0.75, bottom: 0.75,
        header: 0.3, footer: 0.3
    };
    // Automatic page breaks are handled by Excel.
}

/**
 * Auto-fits columns based on the maximum content width in each column, including headers.
 * Considers formatted values for accurate width calculation.
 * @param {ExcelJS.Worksheet} worksheet - The ExcelJS worksheet object.
 */
function autoFitColumns(worksheet) {
    worksheet.columns.forEach(column => {
        // Initialize maxLength with the header text length
        let maxLength = column.header ? column.header.length : 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
            let columnValue = cell.value ? cell.value.toString() : '';
            // If the cell has a number format, try to get the formatted value
            if (cell.numFmt && cell.value instanceof Date) {
                // moment format string might not directly map to Excel's numFmt,
                // but this provides a reasonable estimate for width
                columnValue = moment(cell.value).format('YYYY-MM-DD HH:mm:ss');
            }
            maxLength = Math.max(maxLength, columnValue.length);
        });
        // Set a minimum width and add padding
        column.width = maxLength < 10 ? 10 : maxLength + 2;
    });
}

/**
 * Ensures the export directory exists and then writes the workbook to an Excel file.
 * @param {ExcelJS.Workbook} workbook - The ExcelJS workbook object.
 */
async function exportExcel(workbook) {
    if (!fs.existsSync(EXPORT_DIR)) {
        fs.mkdirSync(EXPORT_DIR, { recursive: true });
        console.log(`Created directory: ${EXPORT_DIR}`);
    }

    // Generate dynamic filename with current date and time
    const timestamp = moment().format('DD MMMM YYYY');
    const dynamicExportFileName = `${timestamp}.xlsx`;
    const dynamicExportFilePath = path.join(EXPORT_DIR, dynamicExportFileName);

    // Write the workbook to the specified file path
    await workbook.xlsx.writeFile(dynamicExportFilePath);
    console.log(`Report successfully exported to ${dynamicExportFilePath}`);
}

// --- Main Execution ---
async function main() {
    try {
        console.log('Loading orders...');
        let orders = loadOrders();
        if (orders.length === 0) {
            console.log('No orders to export. Exiting.');
            return;
        }

        console.log('Sorting orders...');
        // Sort orders: pending orders first, then completed orders.
        // Within each status group, sort by createdAt in ascending order.
        orders.sort((a, b) => {
            const statusOrder = { 'pending': 0, 'done': 1 };
            if (statusOrder[a.status] !== statusOrder[b.status]) {
                return statusOrder[a.status] - statusOrder[b.status];
            }
            return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime();
        });

        console.log('Calculating dashboard statistics...');
        const dashboardData = calculateDashboard(orders);

        console.log('Calculating summary statistics...');
        const { summaryData, totalDrinkQty } = calculateSummary(orders, MENU_ITEMS_CANONICAL);

        console.log('Calculating other items...');
        const otherItemsData = calculateOtherItems(orders, MENU_ITEMS_CANONICAL);

        console.log('Creating Excel workbook...');
        const workbook = new ExcelJS.Workbook();
        buildWorksheet(workbook, orders, dashboardData, summaryData, otherItemsData, MENU_ITEMS_CANONICAL);

        // Get the worksheet to apply auto-fit columns after all data is written
        const worksheet = workbook.getWorksheet('Orders Report');
        autoFitColumns(worksheet);

        console.log('Exporting Excel report...');
        await exportExcel(workbook);

    } catch (error) {
        console.error('An unexpected error occurred during export:', error);
    }
}

// Execute the main function to start the export process
main();
