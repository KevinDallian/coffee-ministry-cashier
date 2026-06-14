# ☕ Café POS — Local LAN Cashier System

A real-time cashier system for food & beverage, running entirely on your local network.
PC handles order entry; tablet handles order status toggling.

---

## Requirements

- [Node.js](https://nodejs.org) v16 or later (installed on your PC)
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
├── server.js          ← Node.js + Express + Socket.io server
├── orders.json        ← Auto-created, persists orders
├── menu.json          ← Auto-created, edit to customize menu
├── package.json
└── public/
    ├── cashier.html   ← PC cashier view
    └── display.html   ← Tablet display view
```
