const express = require("express");
const http = require("http");
const { Server } = require("socket.io");
const fs = require("fs");
const path = require("path");
const { v4: uuidv4 } = require("uuid");
const { exec } = require("child_process"); // Import child_process for executing shell commands

const app = express();
const server = http.createServer(app);
const io = new Server(server);

const DATA_FILE = path.join(__dirname, "orders.json");

// ── Persistence ──────────────────────────────────────────────────────────────
function loadOrders() {
  try {
    if (fs.existsSync(DATA_FILE)) {
      return JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
    }
  } catch (e) {
    console.error("Failed to load orders:", e.message);
  }
  return [];
}

function saveOrders(orders) {
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(orders, null, 2));
  } catch (e) {
    console.error("Failed to save orders:", e.message);
  }
}

let orders = loadOrders();

// ── Default menu ─────────────────────────────────────────────────────────────
const DEFAULT_MENU = [
  { id: "m1", category: "Hot", name: "Hot Latte", price: 2.5 },
  { id: "m2", category: "Hot", name: "Hot Americano", price: 3.0 },
  { id: "m3", category: "Hot", name: "Hot Cappuccino", price: 3.5 },
  { id: "m4", category: "Hot", name: "Espresso", price: 4.0 },
  { id: "m5", category: "Ice", name: "Ice Latte", price: 4.0 },
  { id: "m6", category: "Ice", name: "Ice Americano", price: 4.0 },
  { id: "m7", category: "Kids", name: "Milk", price: 4.0 }
];

const MENU_FILE = path.join(__dirname, "menu.json");
let menu = fs.existsSync(MENU_FILE)
  ? JSON.parse(fs.readFileSync(MENU_FILE))
  : DEFAULT_MENU;

// ── Status flow ───────────────────────────────────────────────────────────────
const STATUS_FLOW = ["pending", "ready", "done"];

// ── Routes ────────────────────────────────────────────────────────────────────
app.use(express.static(path.join(__dirname, "public")));
app.use(express.json());

app.get("/cashier", (req, res) =>
  res.sendFile(path.join(__dirname, "public", "cashier.html"))
);
app.get("/display", (req, res) =>
  res.sendFile(path.join(__dirname, "public", "display.html"))
);

app.get("/api/menu", (req, res) => res.json(menu));
app.get("/api/orders", (req, res) => res.json(orders));

// New endpoint to trigger Excel report generation
app.post("/api/generate-excel", (req, res) => {
  console.log("Received request to generate Excel report.");
  exec("node export_orders.js", (error, stdout, stderr) => {
    if (error) {
      console.error(`exec error: ${error}`);
      console.error(`stderr: ${stderr}`);
      return res.status(500).json({ message: "Failed to generate Excel report.", error: stderr });
    }
    console.log(`stdout: ${stdout}`);
    res.status(200).json({ message: "Excel report generated successfully!", output: stdout });
  });
});

// ── Socket.io ─────────────────────────────────────────────────────────────────
io.on("connection", (socket) => {
  console.log("Client connected:", socket.id);

  // Send current state on connect
  socket.emit("init", { orders, menu });

  // New order from cashier
  socket.on("new_order", (data) => {
    const order = {
      id: uuidv4(),
      tableLabel: data.tableLabel || "Take Away",
      items: data.items,
      total: data.total,
      note: data.note || "",
      status: "pending",
      createdAt: new Date().toISOString(),
    };
    orders.unshift(order);
    saveOrders(orders);
    io.emit("order_added", order);
  });

  // Status toggle from display tablet
  socket.on("update_status", ({ orderId, status }) => {
    const order = orders.find((o) => o.id === orderId);
    if (!order) return;
    order.status = status;
    order.updatedAt = new Date().toISOString();
    saveOrders(orders);
    io.emit("order_updated", order);
  });

  // Advance to next status
  socket.on("next_status", ({ orderId }) => {
    const order = orders.find((o) => o.id === orderId);
    if (!order) return;
    const idx = STATUS_FLOW.indexOf(order.status);
    if (idx < STATUS_FLOW.length - 1) {
      order.status = STATUS_FLOW[idx + 1];
      order.updatedAt = new Date().toISOString();
      saveOrders(orders);
      io.emit("order_updated", order);
    }
  });

  // Delete order
  socket.on("delete_order", ({ orderId }) => {
    orders = orders.filter((o) => o.id !== orderId);
    saveOrders(orders);
    io.emit("order_deleted", { orderId });
  });

  // Clear done orders
  socket.on("clear_done", () => {
    orders = orders.filter((o) => o.status !== "done");
    saveOrders(orders);
    io.emit("orders_cleared", { orders });
  });

  socket.on("disconnect", () => {
    console.log("Client disconnected:", socket.id);
  });
});

// ── Start ─────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
server.listen(PORT, "0.0.0.0", () => {
  const { networkInterfaces } = require("os");
  const nets = networkInterfaces();
  let localIP = "localhost";
  for (const iface of Object.values(nets)) {
    for (const net of iface) {
      if (net.family === "IPv4" && !net.internal) {
        localIP = net.address;
        break;
      }
    }
  }
  const qrcode = require("qrcode-terminal");
  const pcURL = `http://${localIP}:${PORT}/cashier`;
  const tableURL = `http://${localIP}:${PORT}/display`;

  console.clear();

  console.log(`
  ╔══════════════════════════════════════════════════════════════╗
  ║                    ☕ Café Cashier System                    ║
  ╠══════════════════════════════════════════════════════════════╣
  ║  🟢 Server Status : RUNNING                                  ║
  ║  💻 PC Cashier    : ${pcURL.padEnd(36)}     ║
  ║  📱 Tablet View   : ${tableURL.padEnd(36)}     ║
  ╚══════════════════════════════════════════════════════════════╝
  `);

  console.log("┌─────────────────────────────────────────┐");
  console.log("│ 💻 PC Cashier QR Code                   │");
  console.log("└─────────────────────────────────────────┘");
  qrcode.generate(pcURL, { small: true });

  console.log("\n");

  console.log("┌─────────────────────────────────────────┐");
  console.log("│ 📱 Tablet Display QR Code               │");
  console.log("└─────────────────────────────────────────┘");
  qrcode.generate(tableURL, { small: true });

  console.log(`
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Scan the QR code from another device connected to the same Wi-Fi
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  `);
});
