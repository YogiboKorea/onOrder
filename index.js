require('dotenv').config(); 

const express = require('express');
const xlsx = require('xlsx');
const cors = require('cors');
const fs = require('fs');
const bodyParser = require('body-parser');
const { MongoClient, ObjectId } = require('mongodb');

const path = require('path');

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

// --------------------------------------------------------------------------
// 엑셀 -> DB 동기화 함수 (온라인 데이터 전용)
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
        
        console.log(`📄 시트 로드 완료: ${sheetName}`);

        const headerIndex = findHeaderRowIndex(sheet);
        const rawData = xlsx.utils.sheet_to_json(sheet, { range: headerIndex, defval: "" });
        console.log(`📊 엑셀에서 읽은 총 행 개수: ${rawData.length}개`);
        
        if (rawData.length === 0) {
            console.log('⚠️ [경고] 엑셀에 데이터가 없습니다.');
            return 0;
        }

        const firstRow = rawData[0];
        const keys = {
            orderNo: findHeaderKey(firstRow, ['No', 'no', 'NO']),
            date: findHeaderKey(firstRow, ['일자', 'date', 'Date']),
            store: findHeaderKey(firstRow, ['거래처', '지점', '판매처']), // 온라인 거래처
            group1: findHeaderKey(firstRow, ['품목그룹1명', '품목그룹1', '그룹1']), 
            group2: findHeaderKey(firstRow, ['품목그룹2', '그룹2']),
            product: findHeaderKey(firstRow, ['품목명', '규격']),
            qty: findHeaderKey(firstRow, ['수량']),
            amount: findHeaderKey(firstRow, ['합계', '판매금액']), 
            isSet: findHeaderKey(firstRow, ['세트']),
            isCover: findHeaderKey(firstRow, ['커버', '동시'])
        };

        if (!keys.orderNo || !keys.date) {
            console.log('🚨 [치명적 오류] 주문번호나 날짜 컬럼을 찾지 못했습니다.');
            return 0;
        }

        let lastOrderNo = '', lastDate = null, lastStore = '';
        let targetMonths = new Set(); 

        const parsedData = rawData.map((row, idx) => {
            const currentRowIndex = headerIndex + 1 + idx; 

            if (row[keys.orderNo]) lastOrderNo = row[keys.orderNo];
            if (row[keys.date]) lastDate = row[keys.date];
            if (row[keys.store]) lastStore = row[keys.store];

            let amt = 0;
            if (row[keys.amount]) {
                amt = Number(String(row[keys.amount]).replace(/[^0-9.-]+/g, '')) || 0;
            }

            // 오프라인 팝업 등 예외 처리 제거, 단순히 공백만 제거하여 자사몰/스마트스토어 등으로 통합
            let cleanStore = typeof lastStore === 'string' ? lastStore.trim() : lastStore;
            
            const g1 = String(row[keys.group1] || '').trim();
            const g2 = String(row[keys.group2] || '').trim();
            const pName = String(row[keys.product] || '').trim();

            const isNegative = amt < 0; 
            const isExcluded = !isNegative && (pName.includes('쇼핑백') || pName.includes('shopping bag') || g1.includes('부자재'));
            
            if (isExcluded || (!pName && amt === 0)) {
                return null; 
            }

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
                color: color, 
                category: category, 
                beadType: beadType,
                group1: g1,
                group2: g2,
                qty: Number(row[keys.qty]) || 0, 
                amount: amt,
                isSet: (row[keys.isSet] && !String(row[keys.isSet]).includes('해당 없음')),
                isCover: (row[keys.isCover] && !String(row[keys.isCover]).includes('해당 없음'))
            };
        }).filter(d => {
            if (!d) return false;
            if (d.week === '날짜오류') return false;
            if (!d.orderNo) return false;
            return true;
        });

        console.log(`📉 유효 데이터 추출: ${parsedData.length}건`);

        if (parsedData.length === 0) {
            console.log('❌ 유효한 데이터가 없습니다.');
            return 0;
        }

        // 데이터 정리 (세트/커버 여부 주문 단위로 통합)
        const ordersMap = {};
        parsedData.forEach(item => {
            const uniqueKey = item.orderNo; 
            if (!ordersMap[uniqueKey]) {
                ordersMap[uniqueKey] = { hasSet: false, hasCover: false, items: [] };
            }
            ordersMap[uniqueKey].items.push(item);
            if (item.isSet) ordersMap[uniqueKey].hasSet = true;
            if (item.isCover) ordersMap[uniqueKey].hasCover = true;
        });

        const finalOrders = [];
        let finalTotalAmount = 0; 

        Object.values(ordersMap).forEach(order => {
            order.items.forEach(item => {
                item.orderHasSet = order.hasSet;
                item.orderHasCover = order.hasCover;
                finalOrders.push(item);
                finalTotalAmount += item.amount;
            });
        });

        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        const ordersCol = dbOnline.collection('orders'); 

        if (finalOrders.length > 0) {
            const monthArray = Array.from(targetMonths);
            
            if (monthArray.length > 0) {
                console.log(`🧹 기존 데이터 정리 중 (대상 월: ${monthArray.join(', ')})`);
                await ordersCol.deleteMany({ month: { $in: monthArray } });
            }

            const result = await ordersCol.insertMany(finalOrders);
            
            console.log('========================================');
            console.log(`✅ [성공] DB 동기화 완료: ${result.insertedCount}건`);
            console.log(`💰 총 매출 금액: ${finalTotalAmount.toLocaleString()}원`);
            console.log('========================================');

            await dbOnline.collection('system_metadata').updateOne(
                { key: 'last_update_time' },
                { $set: { timestamp: new Date(), updatedCount: finalOrders.length } },
                { upsert: true }
            );
        }

        return finalOrders.length;

    } catch (error) {
        console.error('❌ Excel Sync Error:', error);
        throw error;
    }
}

// ============================================
// API 라우트 정의 (온라인 경로: /api/online/...)
// ============================================

app.post('/api/online/sync', async (req, res) => {
    try {
        const count = await syncExcelToDB();
        res.json({ success: true, message: `${count}건 동기화 완료`, count });
    } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

app.get('/api/online/months', async (req, res) => {
    try {
        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        const months = await dbOnline.collection('orders').distinct('month');
        months.sort().reverse();
        res.json({ success: true, months });
    } catch (err) { res.status(500).json({ success: false, months: [] }); }
});

app.get('/api/online/orders', async (req, res) => {
    try {
        const { month, store } = req.query;
        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        const query = {};
        if (month) query.month = month;
        if (store && store !== 'all') query.store = store; // 여기서 store는 프론트의 channel(자사몰, 스마트스토어 등)
        
        const orders = await dbOnline.collection('orders').find(query).toArray();
        res.json({ success: true, orders });
    } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

app.get('/api/online/system/last-update', async (req, res) => {
    try {
        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        const meta = await dbOnline.collection('system_metadata').findOne({ key: 'last_update_time' });
        if (meta && meta.timestamp) {
            res.json({ success: true, timestamp: meta.timestamp });
        } else {
            res.json({ success: false, message: '기록 없음' });
        }
    } catch (err) { res.status(500).json({ success: false }); }
});


// ============================================
// [신규 추가] 온라인 이벤트 일정 관리 API
// ============================================
const eventsColName = 'events';

// 이벤트 조회 (해당 월에 포함된 이벤트)
app.get('/api/online/events', async (req, res) => {
    try {
        const { month } = req.query; // 예: '2026-03'
        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        // 시작일이나 종료일이 해당 월을 포함하는 데이터 검색
        const query = month ? { $or: [ { startDate: { $regex: `^${month}` } }, { endDate: { $regex: `^${month}` } } ] } : {};
        const events = await dbOnline.collection(eventsColName).find(query).sort({ startDate: 1 }).toArray();
        res.json({ success: true, events });
    } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

// 이벤트 등록
app.post('/api/online/events', async (req, res) => {
    try {
        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        await dbOnline.collection(eventsColName).insertOne(req.body);
        res.json({ success: true });
    } catch (err) { res.status(500).json({ success: false }); }
});

// 이벤트 수정
app.put('/api/online/events/:id', async (req, res) => {
    try {
        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        await dbOnline.collection(eventsColName).updateOne({ _id: new ObjectId(req.params.id) }, { $set: req.body });
        res.json({ success: true });
    } catch (err) { res.status(500).json({ success: false }); }
});

// 이벤트 삭제
app.delete('/api/online/events/:id', async (req, res) => {
    try {
        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        await dbOnline.collection(eventsColName).deleteOne({ _id: new ObjectId(req.params.id) });
        res.json({ success: true });
    } catch (err) { res.status(500).json({ success: false }); }
});


connectDB().then(async () => {
    console.log('⏳ 서버 시작 시 엑셀 초기 동기화 시도...');
    
    // 1. 서버가 켜질 때 DB에 엑셀 데이터를 한번 업데이트 합니다.
    await syncExcelToDB().catch(err => {
        console.log('⚠️ 초기 동기화 실패 (파일 없음 등). 일단 서버는 계속 실행합니다.');
    }); 

    // 2. ★ 제일 중요한 부분! 서버가 꺼지지 않고 계속 대기하도록 만듭니다.
    app.listen(PORT, '0.0.0.0', () => {
        console.log(`🚀 온라인 API 서버가 포트 ${PORT}에서 24시간 정상 대기 중입니다!`);
    });
});