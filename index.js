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
const ONLINE_DB_NAME = 'on';

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

function parseProduct(rawString, group1, memo = '', storeName = '') {
    let namePart = rawString || '-';
    let colorPart = '기타';
    if (typeof rawString === 'string' && rawString.includes('[')) {
        try {
            const parts = rawString.split('[');
            namePart = parts[0].trim();
            if (parts[1]) colorPart = parts[1].replace(']', '').trim();
        } catch (e) { }
    }
    if (namePart.includes('/')) namePart = namePart.split('/')[0].trim();
    if (colorPart.includes('/')) colorPart = colorPart.split('/')[0].trim();

    let category = '기타';
    const lowerName = String(namePart).toLowerCase();
    const lowerRaw = String(rawString || '').toLowerCase();
    const lowerMemo = String(memo || '').toLowerCase();
    const g1 = (group1 || '').trim().toLowerCase();
    const isHomepage = String(storeName).toLowerCase().includes('홈페이지');

    if (lowerRaw.includes('리퍼') || lowerRaw.includes('[리퍼') ||
        (isHomepage && lowerMemo.includes('리퍼'))) {
        category = '리퍼';
    } else if (lowerRaw.includes('한정수량') || lowerRaw.includes('[한정') ||
        lowerRaw.includes('last chance') ||
        lowerMemo.includes('한정수량') || lowerMemo.includes('한정판')) {
        category = '한정수량관';
    } else if (g1.includes('living') || g1.includes('리빙')) {
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
        if (lowerName.includes('소파') || lowerName.includes('맥스') || lowerName.includes('팟') ||
            lowerName.includes('드롭') || lowerName.includes('미디') || lowerName.includes('슬림') ||
            lowerName.includes('더블') || lowerName.includes('라운저') || lowerName.includes('피라미드')) {
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
            isCover: findHeaderKey(firstRow, ['커버', '동시']),
            memo: findHeaderKey(firstRow, ['비고', '메모', 'memo', 'remark', '적요', '특이사항'])
        };

        if (!keys.orderNo || !keys.date) return 0;

        let lastOrderNo = '', lastDate = null, lastStore = '';
        let targetDates = new Set(); // ★ 수정됨: 월(Month) 단위가 아니라 일(Date) 단위로 기록

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
            const memoText = String(row[keys.memo] || '').trim();

            const isNegative = amt < 0;

            const lowerPName = pName.toLowerCase();
            const lowerMemo = memoText.toLowerCase();

            // 배송비 및 불필요 항목 필터링
            const isExcluded = !isNegative && (
                lowerPName.includes('쇼핑백') ||
                lowerPName.includes('shopping bag') ||
                g1.includes('부자재') ||
                lowerPName.includes('배송비') ||
                lowerPName.includes('delivery charge') ||
                lowerMemo.includes('배송비') ||
                lowerMemo.includes('delivery charge')
            );

            if (isExcluded || (!pName && amt === 0)) return null;

            const { name, color, category } = parseProduct(row[keys.product], row[keys.group1], memoText, cleanStore);
            const dInfo = calculateDateInfo(lastDate);

            // ★ 수정됨: 엑셀 파일에 존재하는 날짜들을 모두 수집
            if (dInfo.fullDate !== '-' && dInfo.fullDate !== '날짜오류') {
                targetDates.add(dInfo.fullDate);
            }

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
                isCover: (row[keys.isCover] && !String(row[keys.isCover]).includes('해당 없음')),
                memo: memoText
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
            // ★ 핵심 로직: 엑셀에 포함된 '특정 일자'들만 DB에서 삭제 후 덮어쓰기
            const dateArray = Array.from(targetDates);
            if (dateArray.length > 0) {
                console.log(`🗑 DB에서 삭제 후 덮어쓰기 진행하는 일자: ${dateArray.join(', ')}`);
                await ordersCol.deleteMany({ date: { $in: dateArray } });
            }

            await ordersCol.insertMany(finalOrders);
            await dbOnline.collection('system_metadata').updateOne(
                { key: 'last_update_time' },
                { $set: { timestamp: new Date(), updatedCount: finalOrders.length } },
                { upsert: true }
            );

            const referCount = finalOrders.filter(o => o.category === '리퍼').length;
            const limitedCount = finalOrders.filter(o => o.category === '한정수량관').length;
            if (referCount > 0 || limitedCount > 0) {
                console.log(`📦 특수 카테고리 분류: 리퍼 ${referCount}건, 한정수량관 ${limitedCount}건`);
            }
        }
        return finalOrders.length;
    } catch (error) { throw error; }
}



//기능 추가하하기
// 팝업 스토어 관련 API
app.post('/api/popup/sales', async (req, res) => {
    try {
        const { items, customerName, customerPhone, customerAddress, memo } = req.body;
        // items is an array of { product, color, qty }
        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        const popupCol = dbOnline.collection('popup_sales');
        
        if (!items || !items.length) {
            return res.status(400).json({ success: false, message: '상품이 없습니다.' });
        }

        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        const isoDate = `${year}-${month}-${day}`;

        const newSales = items.map(item => {
            const price = 69000;
            return {
                date: isoDate,
                timestamp: now,
                product: item.product,
                color: item.color,
                qty: Number(item.qty),
                price,
                totalAmount: price * Number(item.qty),
                customerName,
                customerPhone,
                customerAddress,
                memo,
                status: 'SALE'
            };
        });

        const result = await popupCol.insertMany(newSales);
        res.status(200).json({ success: true, message: '판매가 등록되었습니다.', count: result.insertedCount });
    } catch (err) {
        console.error('팝업 판매 등록 에러:', err);
        res.status(500).json({ success: false, message: '서버 에러 발생' });
    }
});

app.post('/api/popup/cancel', async (req, res) => {
    try {
        const { id } = req.body;
        if (!id) return res.status(400).json({ success: false, message: 'ID가 필요합니다.' });

        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        const popupCol = dbOnline.collection('popup_sales');

        const result = await popupCol.updateOne(
            { _id: new ObjectId(id) },
            { $set: { status: 'CANCEL', cancelDate: new Date() } }
        );

        if (result.matchedCount === 0) {
            return res.status(404).json({ success: false, message: '대상을 찾을 수 없습니다.' });
        }
        res.status(200).json({ success: true, message: '취소되었습니다.' });
    } catch (err) {
        console.error('팝업 취소 에러:', err);
        res.status(500).json({ success: false, message: '서버 에러 발생' });
    }
});

app.get('/api/popup/data', async (req, res) => {
    try {
        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        const popupCol = dbOnline.collection('popup_sales');
        const invCol = dbOnline.collection('popup_inventory');

        const sales = await popupCol.find({}).sort({ timestamp: -1 }).toArray();

        // MongoDB에서 설정된 초기 재고 가져오기
        let inventoryData = await invCol.find({}).toArray();
        if (inventoryData.length === 0) {
            // 초기 세팅이 없으면 생성
            const initialSetup = [
                { _id: '마인드 필로우_화이트 미스트', total: 50 },
                { _id: '마인드 필로우_블루문', total: 30 },
                { _id: '마인드 필로우_핑크샌드', total: 50 },
                { _id: '마인드 바디필로우_화이트 미스트', total: 20 },
                { _id: '마인드 바디필로우_블루문', total: 20 },
                { _id: '마인드 바디필로우_핑크샌드', total: 20 }
            ];
            await invCol.insertMany(initialSetup);
            inventoryData = initialSetup;
        }

        const inventory = {};
        inventoryData.forEach(item => {
            inventory[item._id] = { total: item.total, sold: 0 };
        });

        const dailyData = {};
        let totalRevenue = 0;

        sales.forEach(sale => {
            const date = sale.date;
            if (!dailyData[date]) {
                dailyData[date] = { qty: 0, revenue: 0, cancelQty: 0, cancelRevenue: 0 };
            }

            if (sale.status === 'SALE') {
                dailyData[date].qty += sale.qty;
                dailyData[date].revenue += sale.totalAmount;
                totalRevenue += sale.totalAmount;

                const invKey = `${sale.product}_${sale.color}`;
                if (inventory[invKey]) {
                    inventory[invKey].sold += sale.qty;
                }
            } else if (sale.status === 'CANCEL') {
                dailyData[date].cancelQty += sale.qty;
                dailyData[date].cancelRevenue += sale.totalAmount;
                totalRevenue -= sale.totalAmount;

                const invKey = `${sale.product}_${sale.color}`;
                if (inventory[invKey]) {
                    inventory[invKey].sold -= sale.qty;
                }
            }
        });

        res.status(200).json({ success: true, inventory, dailyData, totalRevenue, sales });
    } catch (err) {
        console.error('팝업 데이터 조회 에러:', err);
        res.status(500).json({ success: false, message: '서버 에러 발생' });
    }
});

app.get('/api/popup/excel', async (req, res) => {
    try {
        const dbOnline = mongoClient.db(ONLINE_DB_NAME);
        const popupCol = dbOnline.collection('popup_sales');
        const sales = await popupCol.find({}).sort({ timestamp: -1 }).toArray();

        const dataForExcel = sales.map((sale, index) => ({
            'No': index + 1,
            '판매일자': sale.date,
            '판매시간': new Date(sale.timestamp).toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }),
            '구분': sale.status === 'SALE' ? '판매' : '취소입고',
            '상품명': sale.product,
            '색상': sale.color,
            '수량': sale.qty,
            '단위금액': sale.price,
            '총액': sale.totalAmount,
            '고객명': sale.customerName || '',
            '연락처': sale.customerPhone || '',
            '주소': sale.customerAddress || '',
            '취소일시': sale.cancelDate ? new Date(sale.cancelDate).toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }) : '',
            '메모': sale.memo || ''
        }));

        const wb = xlsx.utils.book_new();
        const ws = xlsx.utils.json_to_sheet(dataForExcel);
        xlsx.utils.book_append_sheet(wb, ws, "현장판매내역");

        const fileBuffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });

        res.setHeader('Content-Disposition', 'attachment; filename="popup_sales_data.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(fileBuffer);
    } catch (err) {
        console.error('팝업 엑셀 다운로드 에러:', err);
        res.status(500).json({ success: false, message: '서버 에러 발생' });
    }
});



// 🚀 실행부: 프로세스 유지 모드로 변경
connectDB().then(async () => {
    console.log('⏳ 엑셀 데이터를 DB로 전송 중입니다...');

    await syncExcelToDB().catch(err => {
        console.log('⚠️ 동기화 실패 (파일 없음 등):', err);
    });

    console.log('✅ DB 전송이 완료되었습니다.');

    app.listen(PORT, () => {
        console.log(`🚀 Server is running on port ${PORT}`);
    });
});