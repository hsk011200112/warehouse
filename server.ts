import express from "express";
import path from "path";
import dotenv from "dotenv";
import { fileURLToPath } from 'url';
import fs from 'fs';
import webpush from 'web-push';

dotenv.config();

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Setup Web Push
const vapidPublicKey = process.env.VAPID_PUBLIC_KEY || 'BKoeHBFmaWtjW78G7FnA_r05KhGaHBbwB3nTqkLN6S79EEZZYGlybW7CvU8Y7Rc7FoMR75OzChA_hGu4VBMHqTM';
const vapidPrivateKey = process.env.VAPID_PRIVATE_KEY || 'jN_lAYwybtQZ1lkHipCWdpjGUt4serT138yvA6q3Q-c';
webpush.setVapidDetails(
  'mailto:your-email@example.com',
  vapidPublicKey,
  vapidPrivateKey
);

let usersCache: any[] | null = null;
let usersCacheTime = 0;
const CACHE_TTL = 60 * 1000; // 60 seconds

async function startServer() {
  const app = express();
  const PORT = process.env.PORT || 3000;

  app.use(express.json({ limit: '50mb' }));

  // Health check endpoint
  app.get("/health", (req, res) => {
    res.status(200).send("OK");
  });

  // VAPID Public Key endpoint for frontend
  app.get("/api/push/vapid-key", (req, res) => {
    res.json({ publicKey: vapidPublicKey });
  });

  // Push Notification trigger endpoint
  app.post("/api/push/send", async (req, res) => {
    const { subscriptions, payload } = req.body;
    
    if (!subscriptions || !Array.isArray(subscriptions)) {
      return res.status(400).json({ error: "Thiếu danh sách subscription" });
    }

    try {
      const stringifiedPayload = JSON.stringify(payload);
      const pushPromises = subscriptions.map((sub: any) => 
        webpush.sendNotification(sub, stringifiedPayload)
          .catch(err => {
            console.error("Lỗi gửi push cho subscription:", err);
            // Có thể trả về lỗi để FE xóa subscription hết hạn nếu statusCode === 410
            return { error: err.statusCode };
          })
      );
      
      const results = await Promise.all(pushPromises);
      res.json({ success: true, results });
    } catch (error: any) {
      console.error("Lỗi Push Notification:", error);
      res.status(500).json({ error: error.message });
    }
  });

  // API Route to authenticate user
  app.post("/api/auth/login", async (req, res) => {
    const { username, password } = req.body;
    const sheetId = process.env.GOOGLE_SHEET_ID;
    const apiKey = process.env.GOOGLE_API_KEY;

    if (!sheetId || !apiKey) {
      return res.status(500).json({ error: "Cấu hình hệ thống bị thiếu" });
    }

    try {
      let users: any[] = [];

      if (usersCache && Date.now() - usersCacheTime < CACHE_TTL) {
        users = usersCache;
      } else {
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/Users!A:E?key=${apiKey}`;
        const response = await fetch(url);
        const data = await response.json();

        if (data.error) {
          if (data.error.message.includes('Rate exceeded') || data.error.message.includes('Quota exceeded')) {
            throw new Error('Hệ thống đang quá tải (Rate exceeded). Vui lòng thử lại sau 1 phút.');
          }
          throw new Error(data.error.message);
        }

        const rows = data.values || [];
        if (rows.length < 2) return res.status(401).json({ error: "Không tìm thấy dữ liệu" });

        const headers = rows[0].map((h: string) => h.toLowerCase().trim());
        users = rows.slice(1).map((row: any[]) => {
          const user: any = {};
          headers.forEach((header: string, index: number) => {
            let key = header;
            if (header === 'tên đăng nhập' || header === 'tài khoản') key = 'username';
            if (header === 'mật khẩu') key = 'password';
            if (header === 'họ tên' || header === 'tên') key = 'name';
            if (header === 'vai trò' || header === 'chức vụ') key = 'role';
            
            user[key] = row[index];
          });
          return user;
        });

        // Update cache
        usersCache = users;
        usersCacheTime = Date.now();
      }

      // Default admin account as fallback
      if (username === 'admin' && password === 'admin123') {
        return res.json({
          id: 'admin',
          username: 'admin',
          name: 'Quản trị viên',
          role: 'quản lí'
        });
      }

      const user = users.find((u: any) => u.username === username && u.password === password);

      if (user) {
        const { password: _, ...userWithoutPassword } = user;
        res.json(userWithoutPassword);
      } else {
        res.status(401).json({ error: "Sai tài khoản hoặc mật khẩu" });
      }
    } catch (error: any) {
      res.status(500).json({ error: error.message });
    }
  });

  // API Route to sync inventory to Google Sheets
  app.post("/api/sync/sheets", async (req, res) => {
    const { inventory } = req.body;
    const sheetId = process.env.GOOGLE_SHEET_ID;
    const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
    const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n');

    if (!sheetId || !clientEmail || !privateKey) {
      return res.status(500).json({ error: "Cấu hình Google Sheets (Service Account) bị thiếu" });
    }

    try {
      const { google } = await import('googleapis');
      const auth = new google.auth.JWT(
        clientEmail,
        undefined,
        privateKey,
        ['https://www.googleapis.com/auth/spreadsheets']
      );

      const sheets = google.sheets({ version: 'v4', auth });

      // Prepare data for Inventory sheet
      const headers = ['ID', 'Tên', 'Danh mục', 'Đơn vị', 'Tồn kho', 'Vị trí', 'Ngưỡng tối thiểu', 'Giá nhập', 'Mô tả'];
      const rows = inventory.map((item: any) => [
        item.id,
        item.name,
        item.category,
        item.unit,
        item.actualStock,
        item.location || '',
        item.minThreshold || 0,
        item.importPrice || 0,
        item.description || ''
      ]);

      const values = [headers, ...rows];

      // Update sheet (assuming a sheet named 'Kho' exists)
      try {
        await sheets.spreadsheets.values.update({
          spreadsheetId: sheetId,
          range: 'Kho!A1',
          valueInputOption: 'RAW',
          requestBody: { values },
        });
      } catch (apiError: any) {
        if (apiError.message) {
          if (apiError.message.includes('Rate exceeded') || apiError.message.includes('Quota exceeded')) {
            throw new Error('Hệ thống Google Sheets đang quá tải (Rate exceeded). Vui lòng đợi 1 phút rồi thử đồng bộ lại.');
          }
          if (apiError.message.includes('missing required authentication credential') || apiError.message.includes('Expected OAuth 2 access token')) {
            throw new Error('Xác thực Service Account thất bại. Vui lòng kiểm tra lại GOOGLE_SERVICE_ACCOUNT_EMAIL và GOOGLE_PRIVATE_KEY trong Cài đặt (Settings).');
          }
          if (apiError.message.includes('invalid_grant')) {
            throw new Error('Chứng chỉ Service Account không hợp lệ (invalid_grant). Vui lòng kiểm tra lại GOOGLE_PRIVATE_KEY.');
          }
        }
        throw apiError;
      }

      res.json({ success: true, message: "Đồng bộ thành công" });
    } catch (error: any) {
      console.error('Sync Error:', error);
      res.status(500).json({ error: error.message });
    }
  });

  const isProd = process.env.NODE_ENV === "production";
  
  // Find the dist directory robustly
  let distPath = path.resolve(process.cwd(), 'dist');
  if (!fs.existsSync(distPath)) {
    distPath = path.resolve(__dirname, 'dist');
    if (!fs.existsSync(distPath)) {
      distPath = path.resolve(__dirname, '../dist');
    }
  }

  if (!isProd) {
    try {
      const { createServer: createViteServer } = await import("vite");
      const vite = await createViteServer({
        server: { middlewareMode: true },
        appType: "spa",
      });
      app.use(vite.middlewares);
    } catch (e) {
      console.warn("Vite failed, falling back to static");
    }
  }

  // Phục vụ tệp tĩnh TRƯỚC khi đăng ký route '*'
  if (fs.existsSync(distPath)) {
    app.use(express.static(distPath));
  }

  app.get('*', (req, res, next) => {
    if (req.path.startsWith('/api') || req.path === '/health') return next();
    const indexPath = path.join(distPath, 'index.html');
    if (fs.existsSync(indexPath)) {
      res.sendFile(indexPath);
    } else {
      res.status(404).send("Vui lòng chạy 'npm run build' trước khi khởi động server.");
    }
  });

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server listening on 0.0.0.0:${PORT} (Mode: ${isProd ? 'Prod' : 'Dev'})`);
  });
}

startServer().catch(err => {
  console.error("Failed to start:", err);
  process.exit(1);
});
