require('dotenv').config();

const express = require('express');
const xlsx = require('xlsx');
const cors = require('cors');
const fs = require('fs');
const bodyParser = require('body-parser');
const { MongoClient, ObjectId } = require('mongodb');
const path = require('path');
const axios = require('axios');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.static('public'));
app.use(bodyParser.json());

const MONGO_URI = process.env.MONGO_URI || 'mongodb://localhost:27017';
const DB_NAME = process.env.DB_NAME || 'on';
const ONLINE_DB_NAME = process.env.ONLINE_DB_NAME || 'on';

let client;
let mongoClient;

const EXCEL_PATH = path.join(__dirname, 'file', 'onOrderData.xlsx');

// ============================
// MongoDB Connect
// ============================
async function connectDB() {
  try {
    client = new MongoClient(MONGO_URI);
    await client.connect();
    mongoClient = client;
    console.log(`✅ MongoDB Connected to database: ${DB_NAME}`);
  } catch (err) {
    console.error('❌ MongoDB Connection Error:', err);
    process.exit(1);
  }
}

// ============================
// Excel Parsing Helpers
// ============================
function findHeaderRowIndex(sheet) {
  if (!sheet['!ref']) return 0;
  const range = xlsx.utils.decode_range(sheet['!ref']);
  for (let R = range.s.r; R <= Math.min(range.e.r, 20); ++R) {
    const rowValues = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cell = sheet[xlsx.utils.encode_cell({ r: R, c: C })];
      if (cell && cell.v) rowValues.push(String(cell.v).trim());
    }
    if (rowValues.join(' ').includes('품목') || rowValues.join(' ').includes('그룹')) return R;
  }
  return 0;
}

function findHeaderKey(row, keywords) {
  if (!row) return null;
  const keys = Object.keys(row);
  return keys.find((k) => {
    const cleanKey = k.replace(/\s+/g, '').replace(/[\(\)\-.]/g, '').toLowerCase();
    return keywords.some((keyword) => cleanKey.includes(keyword.toLowerCase()));
  });
}

function calculateDateInfo(rawDate) {
  if (!rawDate) return { fullDate: '-', month: '-', week: '미지정' };
  let date;

  if (typeof rawDate === 'number') {
    date = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
  } else {
    const dateStr = String(rawDate).trim();
    const match = dateStr.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
    if (match) date = new Date(`${match[1]}-${match[2]}-${match[3]}`);
    else date = new Date(dateStr);
  }

  if (isNaN(date.getTime())) return { fullDate: String(rawDate), month: '-', week: '날짜오류' };

  const year = date.getFullYear();
  const month = date.getMonth();
  const day = date.getDate();
  const mm = String(month + 1).padStart(2, '0');
  const dd = String(day).padStart(2, '0');

  const firstDayOfMonth = new Date(year, month, 1);
  const startDay = firstDayOfMonth.getDay();

  let firstSundayDate;
  if (startDay === 0) firstSundayDate = 1;
  else firstSundayDate = 1 + (7 - startDay);

  let weekNumber;
  if (day <= firstSundayDate) weekNumber = 1;
  else weekNumber = Math.ceil((day - firstSundayDate) / 7) + 1;

  return { fullDate: `${year}-${mm}-${dd}`, month: `${year}-${mm}`, week: `${weekNumber}주차` };
}

function parseProduct(rawString, group1) {
  let namePart = rawString || '-';
  let colorPart = '기타';

  if (typeof rawString === 'string' && rawString.includes('[')) {
    try {
      const parts = rawString.split('[');
      namePart = parts[0].trim();
      if (parts[1]) colorPart = parts[1].replace(']', '').trim();
    } catch (e) {}
  }
  if (namePart.includes('/')) namePart = namePart.split('/')[0].trim();
  if (colorPart.includes('/')) colorPart = colorPart.split('/')[0].trim();

  let category = '기타';
  const g1 = (group1 || '').trim().toLowerCase();

  if (g1.includes('living') || g1.includes('리빙')) category = '리빙';
  else if (g1.includes('sofa') || g1.includes('소파')) category = '소파';
  else if (g1.includes('kids') || g1.includes('키즈')) category = '키즈';
  else if (g1.includes('care') || g1.includes('케어')) category = '케어';
  else if (g1.includes('body') || g1.includes('바디필로우')) category = '바디필로우';
  else {
    const lowerName = String(namePart).toLowerCase();
    if (lowerName.includes('소파') || lowerName.includes('맥스')) category = '소파';
    else if (lowerName.includes('바디필로우') || lowerName.includes('롤')) category = '바디필로우';
    else if (lowerName.includes('인형')) category = '키즈';
    else if (lowerName.includes('리필')) category = '케어';
    else if (lowerName.includes('쿠션')) category = '리빙';
  }

  return { name: namePart, color: colorPart, category };
}

// ============================
// Excel -> DB Sync
// ============================
async function syncExcelToDB() {
  if (!fs.existsSync(EXCEL_PATH)) {
    console.log('❌ [오류] 엑셀 파일이 경로에 없습니다:', EXCEL_PATH);
    return 0;
  }
  console.log(`📂 엑셀 파일 읽기 시작: ${EXCEL_PATH}`);

  const workbook = xlsx.readFile(EXCEL_PATH);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const headerIndex = findHeaderRowIndex(sheet);
  const rawData = xlsx.utils.sheet_to_json(sheet, { range: headerIndex, defval: "" });
  if (rawData.length === 0) return 0;

  const firstRow = rawData[0];
  const keys = {
    orderNo: findHeaderKey(firstRow, ['No', 'no', 'NO']),
    date: findHeaderKey(firstRow, ['일자', 'date', 'Date']),
    store: findHeaderKey(firstRow, ['거래처', '지점', '판매처']),
    group1: findHeaderKey(firstRow, ['품목그룹1명', '품목그룹1', '그룹1']),
    group2: findHeaderKey(firstRow, ['품목그룹2', '그룹2']),
    product: findHeaderKey(firstRow, ['품목명', '규격']),
    qty: findHeaderKey(firstRow, ['수량']),
    amount: findHeaderKey(firstRow, ['합계', '판매금액']),
    isSet: findHeaderKey(firstRow, ['세트']),
    isCover: findHeaderKey(firstRow, ['커버', '동시']),
  };

  if (!keys.orderNo || !keys.date) return 0;

  let lastOrderNo = '';
  let lastDate = null;
  let lastStore = '';
  let targetMonths = new Set();

  const parsedData = rawData
    .map((row, idx) => {
      const currentRowIndex = headerIndex + 1 + idx;

      if (row[keys.orderNo]) lastOrderNo = row[keys.orderNo];
      if (row[keys.date]) lastDate = row[keys.date];
      if (row[keys.store]) lastStore = row[keys.store];

      let amt = Number(String(row[keys.amount]).replace(/[^0-9.-]+/g, '')) || 0;
      let cleanStore = typeof lastStore === 'string' ? lastStore.trim() : lastStore;

      const g1 = String(row[keys.group1] || '').trim();
      const g2 = String(row[keys.group2] || '').trim();
      const pName = String(row[keys.product] || '').trim();

      const isNegative = amt < 0;
      const isExcluded = !isNegative && (pName.includes('쇼핑백') || pName.includes('shopping bag') || g1.includes('부자재'));

      if (isExcluded || (!pName && amt === 0)) return null;

      const { name, color, category } = parseProduct(row[keys.product], row[keys.group1]);
      const dInfo = calculateDateInfo(lastDate);
      if (dInfo.month !== '-') targetMonths.add(dInfo.month);

      let beadType = '기타';
      const g2Lower = g2.toLowerCase();
      if (g2Lower.includes('premium plus')) beadType = 'Premium Plus';
      else if (g2Lower.includes('premium')) beadType = 'Premium';
      else if (g2Lower.includes('standard')) beadType = 'Standard';

      return {
        rowId: currentRowIndex,
        orderNo: lastOrderNo,
        date: dInfo.fullDate,
        month: dInfo.month,
        week: dInfo.week,
        store: cleanStore || '미지정',
        productName: name,
        color,
        category,
        beadType,
        group1: g1,
        group2: g2,
        qty: Number(row[keys.qty]) || 0,
        amount: amt,
        isSet: (row[keys.isSet] && !String(row[keys.isSet]).includes('해당 없음')),
        isCover: (row[keys.isCover] && !String(row[keys.isCover]).includes('해당 없음')),
      };
    })
    .filter((d) => d && d.week !== '날짜오류' && d.orderNo);

  if (parsedData.length === 0) return 0;

  const ordersMap = {};
  parsedData.forEach((item) => {
    const uniqueKey = item.orderNo;
    if (!ordersMap[uniqueKey]) ordersMap[uniqueKey] = { hasSet: false, hasCover: false, items: [] };
    ordersMap[uniqueKey].items.push(item);
    if (item.isSet) ordersMap[uniqueKey].hasSet = true;
    if (item.isCover) ordersMap[uniqueKey].hasCover = true;
  });

  const finalOrders = [];
  Object.values(ordersMap).forEach((order) => {
    order.items.forEach((item) => {
      item.orderHasSet = order.hasSet;
      item.orderHasCover = order.hasCover;
      finalOrders.push(item);
    });
  });

  const dbOnline = mongoClient.db(ONLINE_DB_NAME);
  const ordersCol = dbOnline.collection('orders');

  if (finalOrders.length > 0) {
    const monthArray = Array.from(targetMonths);
    if (monthArray.length > 0) await ordersCol.deleteMany({ month: { $in: monthArray } });
    await ordersCol.insertMany(finalOrders);
    await dbOnline.collection('system_metadata').updateOne(
      { key: 'last_update_time' },
      { $set: { timestamp: new Date(), updatedCount: finalOrders.length } },
      { upsert: true }
    );
  }

  return finalOrders.length;
}

// ============================================
// 📊 기존 매출 데이터 API
// ============================================
app.post('/api/online/sync', async (req, res) => {
  try {
    const count = await syncExcelToDB();
    res.json({ success: true, count });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/online/months', async (req, res) => {
  try {
    const dbOnline = mongoClient.db(ONLINE_DB_NAME);
    const months = await dbOnline.collection('orders').distinct('month');
    months.sort().reverse();
    res.json({ success: true, months });
  } catch (err) {
    res.status(500).json({ success: false, months: [] });
  }
});

app.get('/api/online/orders', async (req, res) => {
  try {
    const { month, store } = req.query;
    const dbOnline = mongoClient.db(ONLINE_DB_NAME);
    const query = {};
    if (month) query.month = month;
    if (store && store !== 'all') query.store = store;
    const orders = await dbOnline.collection('orders').find(query).toArray();
    res.json({ success: true, orders });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/online/system/last-update', async (req, res) => {
  try {
    const meta = await mongoClient.db(ONLINE_DB_NAME).collection('system_metadata').findOne({ key: 'last_update_time' });
    res.json(meta && meta.timestamp ? { success: true, timestamp: meta.timestamp } : { success: false, message: '기록 없음' });
  } catch (err) {
    res.status(500).json({ success: false });
  }
});

// ============================================
// 📅 이벤트 캘린더 관리 API
// ============================================
const eventsColName = 'events';

app.get('/api/online/events', async (req, res) => {
  try {
    const { month } = req.query;
    const dbOnline = mongoClient.db(ONLINE_DB_NAME);
    const query = month
      ? { $or: [{ startDate: { $regex: `^${month}` } }, { endDate: { $regex: `^${month}` } }] }
      : {};
    const events = await dbOnline.collection(eventsColName).find(query).sort({ startDate: 1 }).toArray();
    res.json({ success: true, events });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/online/events', async (req, res) => {
  try {
    await mongoClient.db(ONLINE_DB_NAME).collection(eventsColName).insertOne(req.body);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false });
  }
});

app.put('/api/online/events/:id', async (req, res) => {
  try {
    await mongoClient.db(ONLINE_DB_NAME).collection(eventsColName).updateOne(
      { _id: new ObjectId(req.params.id) },
      { $set: req.body }
    );
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false });
  }
});

app.delete('/api/online/events/:id', async (req, res) => {
  try {
    await mongoClient.db(ONLINE_DB_NAME).collection(eventsColName).deleteOne({ _id: new ObjectId(req.params.id) });
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false });
  }
});

// ============================================
// 🛒 Cafe24 연동 (Proxy + DB 저장)
// ============================================

function getCafe24ConfigOrThrow() {
  const mallId = process.env.CAFE24_MALL_ID;
  const accessToken = process.env.CAFE24_ACCESS_TOKEN;
  const apiVersion = process.env.CAFE24_API_VERSION || '2023-12-01';

  if (!mallId || !accessToken) {
    const err = new Error('서버에 Cafe24 권한 정보가 설정되지 않았습니다.');
    err.statusCode = 400;
    throw err;
  }

  return { mallId, accessToken, apiVersion };
}

// 1) Cafe24에서 상품 즉시 조회(프록시)
app.get('/api/online/cafe24/products', async (req, res) => {
  try {
    const { search, limit, offset, display } = req.query;
    const { mallId, accessToken, apiVersion } = getCafe24ConfigOrThrow();

    const q = new URLSearchParams();
    q.set('display', display || 'T'); // T: 진열중
    if (search) q.set('product_name', search);
    if (limit) q.set('limit', String(limit));
    if (offset) q.set('offset', String(offset));

    const url = `https://${mallId}.cafe24api.com/api/v2/admin/products?${q.toString()}`;

    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'X-Cafe24-Api-Version': apiVersion,
      },
      timeout: 30000,
    });

    res.json({ success: true, products: response.data.products || [], meta: response.data.meta || null });
  } catch (error) {
    const status = error.statusCode || error.response?.status || 500;
    console.error('Cafe24 API 통신 에러:', error.response ? error.response.data : error.message);
    res.status(status).json({ success: false, message: 'Cafe24 상품을 불러오는데 실패했습니다.' });
  }
});

// 2) Cafe24 상품을 DB로 동기화(업서트)
// - pagination: offset/limit 사용
// - 저장 컬렉션: cafe24_products
app.post('/api/online/cafe24/sync-products', async (req, res) => {
  try {
    const { display = 'T', limit = 100, maxPages = 50 } = req.body || {};
    const { mallId, accessToken, apiVersion } = getCafe24ConfigOrThrow();

    const dbOnline = mongoClient.db(ONLINE_DB_NAME);
    const col = dbOnline.collection('cafe24_products');

    // 인덱스(최초 1회만 만들어짐)
    await col.createIndex({ mallId: 1, product_no: 1 }, { unique: true });
    await col.createIndex({ product_name: 'text' });

    let offset = 0;
    let page = 0;
    let totalUpserted = 0;
    let totalFetched = 0;

    while (page < Number(maxPages)) {
      const q = new URLSearchParams();
      q.set('display', display); // T/F/A 등
      q.set('limit', String(limit));
      q.set('offset', String(offset));

      const url = `https://${mallId}.cafe24api.com/api/v2/admin/products?${q.toString()}`;

      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
          'X-Cafe24-Api-Version': apiVersion,
        },
        timeout: 30000,
      });

      const products = response.data.products || [];
      totalFetched += products.length;

      if (products.length === 0) break;

      // bulk upsert
      const ops = products.map((p) => ({
        updateOne: {
          filter: { mallId, product_no: p.product_no },
          update: {
            $set: {
              mallId,
              product_no: p.product_no,
              product_code: p.product_code,
              product_name: p.product_name,
              display: p.display,
              selling: p.selling,
              price: p.price,
              retail_price: p.retail_price,
              supply_price: p.supply_price,
              created_date: p.created_date,
              updated_date: p.updated_date,
              // 원문 전체도 같이 저장(필요하면)
              raw: p,
              syncedAt: new Date(),
            },
          },
          upsert: true,
        },
      }));

      const r = await col.bulkWrite(ops, { ordered: false });
      totalUpserted += (r.upsertedCount || 0) + (r.modifiedCount || 0);

      // 다음 페이지
      offset += Number(limit);
      page += 1;
    }

    await dbOnline.collection('system_metadata').updateOne(
      { key: 'cafe24_products_last_sync' },
      { $set: { timestamp: new Date(), totalFetched, totalUpserted, mallId } },
      { upsert: true }
    );

    res.json({ success: true, totalFetched, totalUpserted });
  } catch (error) {
    const status = error.statusCode || error.response?.status || 500;
    console.error('Cafe24 상품 동기화 에러:', error.response ? error.response.data : error.message);
    res.status(status).json({ success: false, message: 'Cafe24 상품 동기화에 실패했습니다.' });
  }
});

// 3) DB에 저장된 Cafe24 상품 조회/검색
app.get('/api/online/cafe24/db-products', async (req, res) => {
  try {
    const { search, limit = 50, skip = 0, display } = req.query;
    const dbOnline = mongoClient.db(ONLINE_DB_NAME);
    const col = dbOnline.collection('cafe24_products');

    const query = {};
    if (display) query.display = display; // 'T' 등
    if (search) {
      // text index 기반 검색(대안: product_name regex)
      query.$text = { $search: String(search) };
    }

    const cursor = col.find(query, search ? { projection: { score: { $meta: 'textScore' } } } : {});
    if (search) cursor.sort({ score: { $meta: 'textScore' } });
    else cursor.sort({ updated_date: -1 });

    const items = await cursor.skip(Number(skip)).limit(Number(limit)).toArray();
    res.json({ success: true, products: items });
  } catch (err) {
    res.status(500).json({ success: false, message: 'DB 상품 조회에 실패했습니다.' });
  }
});

app.get('/api/online/cafe24/last-sync', async (req, res) => {
  try {
    const meta = await mongoClient.db(ONLINE_DB_NAME).collection('system_metadata').findOne({ key: 'cafe24_products_last_sync' });
    res.json(meta?.timestamp ? { success: true, timestamp: meta.timestamp, meta } : { success: false, message: '기록 없음' });
  } catch (err) {
    res.status(500).json({ success: false });
  }
});

// ============================
// Server Start
// ============================
connectDB().then(async () => {
  console.log('⏳ 서버 시작 시 엑셀 초기 동기화 시도...');
  await syncExcelToDB().catch(() => {
    console.log('⚠️ 초기 동기화 실패 (파일 없음 등). 일단 서버는 계속 실행합니다.');
  });

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 온라인 API 서버가 포트 ${PORT}에서 24시간 정상 대기 중입니다!`);
  });
});
