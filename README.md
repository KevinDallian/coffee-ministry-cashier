# ☕ Café POS — Local LAN Cashier System

A real-time cashier system for food & beverage, running entirely on your local network.
PC handles order entry; tablet handles order status toggling.

---

## Requirements

- [Node.js](https://nodejs.org) **v12.0.0** or higher (Note: Node.js 12 reached End-of-Life in April 2022 and is no longer recommended due to security risks.)
- Supports **Windows 7 32-bit** (via Node.js v12)
- PC and tablet connected to the **same Wi-Fi network**

---

## Setup

```bash
# 1. Install dependencies (only needed once)
npm install

# 2. Start the server
node server.js
```

When the server starts, it prints the URLs:

```
╔══════════════════════════════════════════╗
║       ☕  Café Cashier System             ║
╠══════════════════════════════════════════╣
║  PC Cashier  → http://192.168.1.X:3000/cashier
║  Tablet View → http://192.168.1.X:3000/display
╚══════════════════════════════════════════╝
```

- **PC**: Open the Cashier URL in your browser
- **Tablet**: Open the Display URL in your browser (or scan a QR code)

---

## Features

### Cashier View (PC)
- Browse menu by category (Coffee, Tea, Food, Cold)
- Tap items to add to order
- Adjust quantities with +/−
- Select table (Take Away, T1–T6)
- Add order notes
- Place order — tablet sees it instantly

### Display View (Tablet)
- Real-time order cards appear as orders are placed
- Filter by: All / Pending / Preparing / Ready / Done
- Tap **▶ Start Preparing** → moves to Preparing
- Tap **✓ Mark Ready** → moves to Ready
- Tap **✓ Done** → marks complete
- Clear all Done orders with one tap

---

## Excel Report Generation

This system includes a script (`export_orders.js`) to generate a detailed Excel (.xlsx) report of all orders.

### Features:
-   **Dynamic Filename**: Reports are saved with a timestamp in the format "DD MMMM YYYY - HH:mm - orders-report.xlsx".
-   **Comprehensive Data**: Includes a dashboard summary, a main table of all orders, and detailed summary sections for menu items and other/unknown items.
-   **Themed Formatting**:
    -   Header Fill: Dark Green (`#264e13`) with White, bold Arial font.
    -   Alternating data rows: Light Gray (`#e2e7df`) and White with Arial font.
    -   Columns auto-fit to content and header length.
    -   `Created At` and `Updated At` columns display time only (`hh:mm:ss`).
    -   Summary and "Other Items" tables are positioned starting at column `O`.
-   **Print-Friendly Layout**: Configured for landscape orientation, fit to page width, repeating header row, and A4 margins.

### How to Generate:
1.  On the **Cashier View (PC)**, click the "**Generate Excel Report**" button in the header.
2.  Confirm the prompt.
3.  The report will be saved to the `./exports/` directory in your project root.

---

## Customizing the Menu

Edit `menu.json` (created after first run) or modify `DEFAULT_MENU` in `server.js`:

```json
[
  { "id": "m1", "category": "Coffee", "name": "Espresso", "price": 2.50 },
  ...
]
```

---

## Data

Orders are saved to `orders.json` automatically and survive server restarts.
To wipe all orders, delete `orders.json` and restart.

---

## Files

```
cashier/
├── .gitignore         ← Specifies intentionally untracked files
├── server.js          ← Node.js + Express + Socket.io server
├── export_orders.js   ← Script to generate Excel reports
├── orders.json        ← Auto-created, persists orders (ignored by Git)
├── menu.json          ← Auto-created, edit to customize menu
├── package.json
├── package-lock.json  ← Dependency lock file (ignored by Git)
├── public/
│   ├── cashier.html   ← PC cashier view
│   └── display.html   ← Tablet display view
└── exports/           ← Directory for generated Excel reports (ignored by Git)
```
