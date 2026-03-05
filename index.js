require('dotenv').config(); 

const express = require('express');
const xlsx = require('xlsx');
const cors = require('cors');
const fs = require('fs');
const bodyParser = require('body-parser');
const { MongoClient, ObjectId } = require('mongodb'); // ObjectId 필수!
const path = require('path');
const axios = require('axios'); // ★ Cafe24 연동을 위해 axios 추가

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.static('public')); 
app.use(bodyParser.json());

const MONGO_URI = process.env.MONGO_URI || 'mongodb://localhost:27017'; 
const DB_NAME = process.env.DB_NAME || 'on';
const ONLINE_DB_NAME = 'on'; // 온라인 데이터를 저장할 DB명

let client;
let mongoClient; 

const EXCEL_PATH = path.join(__dirname, 'file', 'onOrderData.xlsx');

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

function findHeaderRowIndex(sheet) {
    if (!sheet['!ref']) return 0;
    const range = xlsx.utils.decode_range(sheet['!ref']);
    for (let R = range.s.r; R <= Math.min(range.e.r, 20); ++R) {
        const rowValues = [];
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cell = sheet[xlsx.utils.encode_cell({r: R, c: C})];
            if (cell && cell.v) rowValues.push(String(cell.v).trim());
        }
        if (rowValues.join(' ').includes('품목') || rowValues.join(' ').includes('그룹')) return R;
    }
    return 0;
}

function findHeaderKey(row, keywords) {
    if (!row) return null;
    const keys = Object.keys(row);
    return keys.find(k => {
        const cleanKey = k.replace(/\s+/g, '').replace(/[\(\)\-.]/g, '').toLowerCase(); 
        return keywords.some(keyword => cleanKey.includes(keyword));
    });
}

function calculateDateInfo(rawDate) {
    if (!rawDate) return { fullDate: '-', month: '-', week: '미지정' };
    let date;
    if (typeof rawDate === 'number') { date = new Date(Math.round((rawDate - 25569) * 86400 * 1000)); } 
    else {
        const dateStr = String(rawDate).trim();
        const match = dateStr.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
        if (match) date = new Date(`${match[1]}-${match[2]}-${match[3]}`); else date = new Date(dateStr);
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
    if (day <= firstSundayDate) { weekNumber = 1; } 
    else { weekNumber = Math.ceil((day - firstSundayDate) / 7) + 1; }
    const weekLabel = weekNumber + '주차';

    return { fullDate: `${year}-${mm}-${dd}`, month: `${year}-${mm}`, week: weekLabel };
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
    const lowerName = String(namePart).toLowerCase();
    const lowerRaw = String(rawString || '').toLowerCase();
    const g1 = (group1 || '').trim().toLowerCase(); 

    // ★ 1순위: 리퍼 / 한정수량관 (상품명 기준 우선 분류)
    if (lowerRaw.includes('리퍼') || lowerRaw.includes('[리퍼')) {
        category = '리퍼';
    } else if (lowerRaw.includes('한정수량') || lowerRaw.includes('[한정') || lowerRaw.includes('last chance')) {
        category = '한정수량관';
    }
    // ★ 2순위: 기존 카테고리 분류 (리퍼/한정이 아닌 경우)
    else if (g1.includes('living') || g1.includes('리빙')) {
        category = '리빙';
    } else if (g1.includes('sofa') || g1.includes('소파')) {
        category = '소파';
    } else if (g1.includes('kids') || g1.includes('키즈')) {
        category = '키즈';
    } else if (g1.includes('care') || g1.includes('케어')) {
        category = '케어';
    } else if (g1.includes('body') || g1.includes('바디필로우')) {
        category = '바디필로우';
    } else if (g1.includes('cover') || g1.includes('커버')) {
        category = '커버';
    } else {
        // 상품명 기반 추가 분류
        if (lowerName.includes('소파') || lowerName.includes('맥스') || lowerName.includes('팟') || 
            lowerName.includes('드롭') || lowerName.includes('미디') || lowerName.includes('슬림')) {
            category = '소파';
        } else if (lowerName.includes('바디필로우') || lowerName.includes('롤') || lowerName.includes('캐터필러')) {
            category = '바디필로우';
        } else if (lowerName.includes('커버')) {
            category = '커버';
        } else if (lowerName.includes('인형') || lowerName.includes('메이트')) {
            category = '키즈';
        } else if (lowerName.includes('리필') || lowerName.includes('보충재')) {
            category = '케어';
        } else if (lowerName.includes('쿠션') || lowerName.includes('블랭킷') || lowerName.includes('토퍼')) {
            category = '리빙';
        }
    }
    
    return { name: namePart, color: colorPart, category };
}

// 주문 메모 업데이트 API
app.patch('/api/ordersOffData/:id/memo', async (req, res) => {
    try {
    const { id } = req.params;
    const { cs_memo } = req.body;
    
    // ✅ 수정됨: mongoClient를 이용해 DB와 컬렉션을 정확히 지정
    const dbOnline = mongoClient.db(ONLINE_DB_NAME);
    const result = await dbOnline.collection('ordersOffData').updateOne(
    { _id: new ObjectId(id) },
    { $set: { cs_memo: cs_memo } }
    );
    
    res.json({ success: true, message: '메모가 업데이트되었습니다.' });
    } catch (error) {
    console.error('메모 업데이트 오류:', error);
    res.status(500).json({ success: false, message: '서버 오류' });
    }
});

// --------------------------------------------------------------------------
// 엑셀 -> DB 동기화 함수
// --------------------------------------------------------------------------
async function syncExcelToDB() {
    if (!fs.existsSync(EXCEL_PATH)) {
        console.log('❌ [오류] 엑셀 파일이 경로에 없습니다:', EXCEL_PATH);
        return 0;
    }
    console.log(`📂 엑셀 파일 읽기 시작: ${EXCEL_PATH}`);

    try {
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
            isCover: findHeaderKey(firstRow, ['커버', '동시'])
        };

        if (!keys.orderNo || !keys.date) return 0;

        let lastOrderNo = '', lastDate = null, lastStore = '';
        let targetMonths = new Set(); 

        const parsedData = rawData.map((row, idx) => {
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
                rowId: currentRowIndex, orderNo: lastOrderNo, date: dInfo.fullDate, 
                month: dInfo.month, week: dInfo.week, store: cleanStore || '미지정', 
                productName: name, color: color, category: category, 
                beadType: beadType, group1: g1, group2: g2, 
                qty: Number(row[keys.qty]) || 0, amount: amt,
                isSet: (row[keys.isSet] && !String(row[keys.isSet]).includes('해당 없음')),
                isCover: (row[keys.isCover] && !String(row[keys.isCover]).includes('해당 없음'))
            };
        }).filter(d => d && d.week !== '날짜오류' && d.orderNo);

        if (parsedData.length === 0) return 0;

        const ordersMap = {};
        parsedData.forEach(item => {
            const uniqueKey = item.orderNo; 
            if (!ordersMap[uniqueKey]) ordersMap[uniqueKey] = { hasSet: false, hasCover: false, items: [] };
            ordersMap[uniqueKey].items.push(item);
            if (item.isSet) ordersMap[uniqueKey].hasSet = true;
            if (item.isCover) ordersMap[uniqueKey].hasCover = true;
        });

        const finalOrders = [];
        Object.values(ordersMap).forEach(order => {
            order.items.forEach(item => {
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
    } catch (error) { throw error; }
}

// ============================================
// 📊 기존 매출 데이터 API
// ============================================
app.post('/api/online/sync', async (req, res) => {
    try { const count = await syncExcelToDB(); res.json({ success: true, count }); } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

app.get('/api/online/months', async (req, res) => {
    try {
        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        const months = await dbOnline.collection('orders').distinct('month');
        months.sort().reverse(); res.json({ success: true, months });
    } catch (err) { res.status(500).json({ success: false, months: [] }); }
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
    } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

app.get('/api/online/system/last-update', async (req, res) => {
    try {
        const meta = await mongoClient.db(ONLINE_DB_NAME).collection('system_metadata').findOne({ key: 'last_update_time' });
        res.json(meta && meta.timestamp ? { success: true, timestamp: meta.timestamp } : { success: false, message: '기록 없음' });
    } catch (err) { res.status(500).json({ success: false }); }
});

// ============================================
// 📅 이벤트 캘린더 관리 API
// ============================================
const eventsColName = 'events';

app.get('/api/online/events', async (req, res) => {
    try {
        const { month } = req.query; 
        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        const query = month ? { $or: [ { startDate: { $regex: `^${month}` } }, { endDate: { $regex: `^${month}` } } ] } : {};
        const events = await dbOnline.collection(eventsColName).find(query).sort({ startDate: 1 }).toArray();
        res.json({ success: true, events });
    } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

app.post('/api/online/events', async (req, res) => {
    try {
        await mongoClient.db(ONLINE_DB_NAME).collection(eventsColName).insertOne(req.body);
        res.json({ success: true });
    } catch (err) { res.status(500).json({ success: false }); }
});

app.put('/api/online/events/:id', async (req, res) => {
    try {
        await mongoClient.db(ONLINE_DB_NAME).collection(eventsColName).updateOne({ _id: new ObjectId(req.params.id) }, { $set: req.body });
        res.json({ success: true });
    } catch (err) { res.status(500).json({ success: false }); }
});

app.delete('/api/online/events/:id', async (req, res) => {
    try {
        await mongoClient.db(ONLINE_DB_NAME).collection(eventsColName).deleteOne({ _id: new ObjectId(req.params.id) });
        res.json({ success: true });
    } catch (err) { res.status(500).json({ success: false }); }
});

// ============================================
// 🛒 [신규] Cafe24 상품 검색 연동 API (Proxy)
// ============================================
app.get('/api/online/cafe24/products', async (req, res) => {
    try {
        const { search } = req.query; 
        
        // 환경변수에서 카페24 정보 가져오기 (클라우드타입 환경변수 설정 필수!)
        const mallId = process.env.CAFE24_MALL_ID;
        const accessToken = process.env.CAFE24_ACCESS_TOKEN;

        if (!mallId || !accessToken) {
            return res.status(400).json({ 
                success: false, 
                message: '서버에 Cafe24 권한 정보가 설정되지 않았습니다.' 
            });
        }

        // Cafe24 API v2 (상품 목록 검색)
        // 참고: display=T (진열중인 상품만)
        let url = `https://${mallId}.cafe24api.com/api/v2/admin/products?display=T`;
        
        if (search) {
            url += `&product_name=${encodeURIComponent(search)}`;
        }

        // Cafe24로 요청 보내기
        const response = await axios.get(url, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
                'X-Cafe24-Api-Version': '2023-12-01' // 카페24 최신 API 버전
            }
        });

        // 프론트엔드로 데이터 전달
        res.json({ success: true, products: response.data.products });
        
    } catch (error) {
        console.error('Cafe24 API 통신 에러:', error.response ? error.response.data : error.message);
        res.status(500).json({ success: false, message: 'Cafe24 상품을 불러오는데 실패했습니다.' });
    }
});


// 서버 시작
connectDB().then(async () => {
    console.log('⏳ 서버 시작 시 엑셀 초기 동기화 시도...');
    await syncExcelToDB().catch(err => {
        console.log('⚠️ 초기 동기화 실패 (파일 없음 등). 일단 서버는 계속 실행합니다.');
    }); 

    app.listen(PORT, '0.0.0.0', () => {
        console.log(`🚀 온라인 API 서버가 포트 ${PORT}에서 24시간 정상 대기 중입니다!`);
    });
});