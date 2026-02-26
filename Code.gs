// ============================================================
// AUTRANS KPI REPORTING SYSTEM — Google Apps Script v5
// ============================================================
// Layout Excel จริง (จากรูป):
//
//  Row 1: "AUTRANS DELIVERY TRIP CONTROL SHEET (S/R, HINO)"
//  Row 2: (ว่าง)
//  Row 3: |        |       |       | <── Mar'2025 ──>        |TOTAL|PRICE|TOTAL|PRICE|TOTAL| <── Apr'2025 ──>     |TOTAL|PRICE|TOTAL|PRICE|TOTAL
//  Row 4: |PRODUCTS|ROUTE  |SHIFT  |21|22|23|24|25|26|27|..31|TRIP |     |(ATH)|     |(MSP)|1|2|3|4|...|19|20   |TRIP |     |(ATH)|     |(MSP)
//  Row 5: |FRAME...|SRF1   |DAY    |38|31|37|32|41|37|  |   .|216  |1425 |307800|1141|246456|39|39|39|37|...|   |330  |1445 |476850|1157|381810
//  Row 6: |        |       |NIGHT  |39|38|37|33|5 |  |  |   .|152  |1425 |216600|1141|173432|39|39|40|39|...|   |325  |1445 |469625|1157|376025
//  Row 7: |        |       |EXTRA  |  |  |  |  |  |  |  |   .| 0   |1425 |     0|1141|0     |39|...|   |   |1  |271  |1445 |391595|1157|313547
//  Row 8: |        |TOTAL  |       |77|69| -|74|65|46|37| - .|368  |     |524400|    |419888|346|..|80|76|...|   |926  |     |1338070|  |1071382
//  Row 9: |SERVICE.|SERVICE|DAY    |3 |3 |3 |3 |3 |2 |3 |3  |23   |915  |21045 |786 |18078 |3 |3 |3 |3 |...|   |...
// ...
//
// กลยุทธ์:
//  1. หา header row (มี SHIFT + TOTALTRIP)
//  2. นับ TOTALTRIP ทุก index → block1Col = index แรก, block2Col = index ที่สอง
//  3. หา PRICE[0], TOTALATH[0] = block1; PRICE[1], TOTALATH[1] = block2
//  4. date cols: ก่อน block1Col = prev period (21-31); ระหว่าง block1Col+1..block2Col-1 = curr (1-20)
//  5. parse ทุก data row แยก shift (DAY/NIGHT/EXTRA) → บันทึก 2 records ต่อ row (prev + curr)
//  6. ข้ามแถว TOTAL (ROUTE="TOTAL" หรือ PRODUCTS="TOTAL")
//  7. merge cell: ถ้า PRODUCTS หรือ ROUTE ว่าง ใช้ค่า lastProd/lastRoute
// ============================================================

var CONFIG = {
  SHEET_NAMES: ["SR","new DECK","P>BANPHO (7.4 M.)","P>BANPHO (8.4 M.)","SR OEM","BANPHO OEM"],
  DATA_SHEET:    "KPI_DATA",
  SUMMARY_SHEET: "KPI_SUMMARY",
  LOG_SHEET:     "UPLOAD_LOG",
};

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚛 AUTRANS KPI")
    .addItem("📊 เปิด Dashboard",          "openDashboard")
    .addSeparator()
    .addItem("📤 อัปโหลดไฟล์ Excel ใหม่", "showUploadSidebar")
    .addItem("🔄 คำนวณ KPI ใหม่",          "recalculateAll")
    .addItem("📋 ตรวจสอบ Data Quality",    "runQualityCheck")
    .addSeparator()
    .addItem("📥 Export Report Excel",      "exportToExcel")
    .addToUi();
}

function openDashboard() {
  var html = HtmlService.createTemplateFromFile("Dashboard")
    .evaluate().setTitle("AUTRANS KPI Dashboard").setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModelessDialog(html, "📊 AUTRANS KPI Dashboard");
}

function showUploadSidebar() {
  var html = HtmlService.createTemplateFromFile("Sidebar")
    .evaluate().setTitle("อัปโหลดไฟล์ Excel");
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================
// SECTION 1: NORMALIZE + PERIOD HELPERS
// ============================================================

function normCol(s) {
  // ตัด space, _, -, (, ), ., ' แล้ว uppercase
  return String(s).toUpperCase().replace(/[\s_()\-.,\']/g, "");
}

var MON_MAP = {JAN:1,FEB:2,MAR:3,APR:4,MAY:5,JUN:6,JUL:7,AUG:8,SEP:9,OCT:10,NOV:11,DEC:12};

function parsePeriodLabel(cell) {
  // แปลง "Mar'2025" / "APR2025" / "2025-04" / "Apr'25" → "2025-04"
  var s = String(cell).toUpperCase().replace(/['\s]/g, "");
  var m1 = s.match(/^([A-Z]{3})(\d{4})$/);
  if (m1 && MON_MAP[m1[1]]) return m1[2] + "-" + String(MON_MAP[m1[1]]).padStart(2,"0");
  var m2 = s.match(/^([A-Z]{3})(\d{2})$/);  // APR25
  if (m2 && MON_MAP[m2[1]]) return "20" + m2[2] + "-" + String(MON_MAP[m2[1]]).padStart(2,"0");
  var m3 = s.match(/(\d{4})[-_]?(\d{2})/);
  if (m3) return m3[1] + "-" + m3[2];
  return null;
}

function prevPeriod(yyyymm) {
  var p = yyyymm.split("-"); var y = parseInt(p[0]); var m = parseInt(p[1]);
  return m === 1 ? (y-1)+"-12" : y+"-"+String(m-1).padStart(2,"0");
}

// ============================================================
// SECTION 2: MAIN ENTRY POINT (รับ JSON จาก SheetJS)
// ============================================================

function receiveSheetData(payload) {
  try {
    var filename    = payload.filename || "unknown.xlsx";
    var company     = detectCompany(filename);
    var fallback    = detectPeriodMonth(filename);  // period จากชื่อไฟล์ = curr period
    var sheetsData  = payload.sheets || {};

    var allRecords  = [];
    var validations = [];

    CONFIG.SHEET_NAMES.forEach(function(sheetName) {
      var rawRows = sheetsData[sheetName];
      if (!rawRows || rawRows.length < 3) {
        validations.push({ sheet: sheetName, status: "error", emoji: "❌",
          issues: [rawRows ? "Sheet มีข้อมูลน้อยเกินไป" : 'ไม่พบ Sheet "'+sheetName+'"'] });
        return;
      }
      var result = parseAutransSheet(rawRows, sheetName, company, fallback);
      validations.push(result.validation);
      if (result.validation.status !== "error") {
        result.records.forEach(function(r){ allRecords.push(r); });
      }
    });

    if (allRecords.length > 0) saveRecords(allRecords);
    recalculateSilent();
    logUpload(filename, company, fallback, validations);

    return {
      success: true, company: company, periodMonth: fallback,
      recordsStored: allRecords.length, validations: validations,
    };
  } catch(e) {
    Logger.log("receiveSheetData ERROR: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// SECTION 3: CORE PARSER — รองรับ layout AUTRANS จริง
// ============================================================

function parseAutransSheet(rawRows, sheetName, company, fallbackCurrPeriod) {
  var issues = [];

  // ── Step 1: หา header row ──────────────────────────────────
  // header row ต้องมี SHIFT และ TOTALTRIP (normalized)
  var headerIdx = -1;
  for (var i = 0; i < Math.min(10, rawRows.length); i++) {
    var hasShift = rawRows[i].some(function(c){ return normCol(c) === "SHIFT"; });
    var hasTrip  = rawRows[i].some(function(c){ return normCol(c) === "TOTALTRIP"; });
    if (hasShift && hasTrip) { headerIdx = i; break; }
  }

  if (headerIdx < 0) {
    // แสดง normalized ของ 5 แถวแรกเพื่อ debug
    var dbg = rawRows.slice(0, Math.min(5, rawRows.length)).map(function(r, ri) {
      return "R"+(ri+1)+": [" + r.map(function(c){ return normCol(c)||"_"; }).filter(function(v,j){ return j < 10; }).join("|") + "]";
    }).join("  ");
    return { records: [], validation: { sheet: sheetName, status: "error", emoji: "❌",
      issues: ["header row ไม่พบ (ต้องมี SHIFT + TOTALTRIP)", dbg] } };
  }

  var headers = rawRows[headerIdx];

  // ── Step 2: หาตำแหน่ง PRODUCTS, ROUTE, SHIFT ───────────────
  var colPROD=-1, colROUTE=-1, colSHIFT=-1;
  headers.forEach(function(h, i) {
    var n = normCol(h);
    if (n==="PRODUCTS" && colPROD  < 0) colPROD  = i;
    if (n==="ROUTE"    && colROUTE < 0) colROUTE = i;
    if (n==="SHIFT"    && colSHIFT < 0) colSHIFT = i;
  });

  if (colROUTE < 0 || colSHIFT < 0) {
    var hSample = headers.map(function(h){ return normCol(h)||"_"; }).slice(0,15).join("|");
    return { records: [], validation: { sheet: sheetName, status: "error", emoji: "❌",
      issues: ["ไม่พบ ROUTE หรือ SHIFT ใน header row "+headerIdx, "norm: "+hSample] } };
  }

  // ── Step 3: หาตำแหน่ง TOTALTRIP ทั้งหมด ────────────────────
  // block1TripCol = ตำแหน่งแรก (ฝั่ง prev period 21-31)
  // block2TripCol = ตำแหน่งที่สอง (ฝั่ง curr period 1-20)
  var tripCols = [];
  headers.forEach(function(h, i) {
    if (normCol(h) === "TOTALTRIP") tripCols.push(i);
  });

  if (tripCols.length === 0) {
    var normAll = headers.map(function(h){ return normCol(h)||"_"; }).join("|");
    return { records: [], validation: { sheet: sheetName, status: "error", emoji: "❌",
      issues: ["ไม่พบ TOTALTRIP ใน header", "normalized headers: " + normAll.substring(0,400)] } };
  }

  var b1TripCol = tripCols[0];
  var b2TripCol = tripCols.length > 1 ? tripCols[1] : -1;

  // ── Step 4: หา PRICE และ TOTAL(ATH) ของแต่ละ block ─────────
  // เพราะ header มีชื่อซ้ำกัน จึงหาตามลำดับ
  var priceCols = [], athCols = [];
  headers.forEach(function(h, i) {
    var n = normCol(h);
    if (n === "PRICE")    priceCols.push(i);
    if (n === "TOTALATH") athCols.push(i);
  });

  // หา PRICE/ATH ที่อยู่ใกล้หลัง TripCol แต่ละ block
  function findAfter(colList, afterCol, beforeCol) {
    for (var k = 0; k < colList.length; k++) {
      var c = colList[k];
      if (c > afterCol && (beforeCol < 0 || c < beforeCol)) return c;
    }
    return -1;
  }

  var b1PriceCol = findAfter(priceCols, colSHIFT, b2TripCol < 0 ? 9999 : b2TripCol);
  var b1AthCol   = findAfter(athCols,   colSHIFT, b2TripCol < 0 ? 9999 : b2TripCol);
  var b2PriceCol = b2TripCol >= 0 ? findAfter(priceCols, b2TripCol, 9999) : -1;
  var b2AthCol   = b2TripCol >= 0 ? findAfter(athCols,   b2TripCol, 9999) : -1;

  // fallback ถ้าหาไม่ได้
  if (b1PriceCol < 0 && priceCols.length > 0) b1PriceCol = priceCols[0];
  if (b1AthCol   < 0 && athCols.length   > 0) b1AthCol   = athCols[0];
  if (b2PriceCol < 0 && priceCols.length > 1) b2PriceCol = priceCols[1];
  if (b2AthCol   < 0 && athCols.length   > 1) b2AthCol   = athCols[1];

  // ── Step 5: หา date columns แต่ละ block ────────────────────
  var b1DateCols = [], b2DateCols = [];
  headers.forEach(function(h, i) {
    var n = parseInt(String(h).trim());
    if (isNaN(n) || n < 1 || n > 31) return;
    if (i < b1TripCol) b1DateCols.push(i);
    else if (b2TripCol > 0 && i > b1TripCol && i < b2TripCol) b2DateCols.push(i);
  });

  // ── Step 6: ระบุ period จาก month-label row ─────────────────
  // สแกน rows เหนือ headerIdx
  var currPeriod = fallbackCurrPeriod;
  var prevPer    = prevPeriod(currPeriod);

  for (var ri = 0; ri < headerIdx; ri++) {
    var labelRow = rawRows[ri];
    var found = [];
    labelRow.forEach(function(cell) {
      var p = parsePeriodLabel(cell);
      if (p) {
        // ป้องกัน duplicate
        if (found.indexOf(p) < 0) found.push(p);
      }
    });
    if (found.length >= 2) {
      prevPer    = found[0];
      currPeriod = found[1];
      break;
    } else if (found.length === 1) {
      currPeriod = found[0];
      prevPer    = prevPeriod(currPeriod);
      break;
    }
  }

  // ── Step 7: Parse data rows ──────────────────────────────────
  var records    = [];
  var nullRoutes = 0;
  var lastProd   = "";
  var lastRoute  = "";

  for (var r = headerIdx + 1; r < rawRows.length; r++) {
    var row = rawRows[r];

    // ข้ามแถวว่าง
    var allEmpty = row.every(function(c){ var s=String(c).trim(); return s===""||s==="-"||s==="–"; });
    if (allEmpty) continue;

    // อ่าน PRODUCTS, ROUTE, SHIFT
    var prod  = colPROD  >= 0 ? String(row[colPROD]  != null ? row[colPROD]  : "").trim() : "";
    var route = String(row[colROUTE] != null ? row[colROUTE] : "").trim();
    var shift = String(row[colSHIFT] != null ? row[colSHIFT] : "").trim();

    // Merge cell handling
    if (prod)  lastProd  = prod;
    if (route) lastRoute = route;

    // ข้ามแถว TOTAL
    var rNorm = normCol(lastRoute);
    var sNorm = normCol(shift);
    if (rNorm === "TOTAL" || sNorm === "TOTAL") continue;

    // ถ้า ROUTE ว่างและ SHIFT ว่าง → ข้าม
    if (!lastRoute && !shift) { nullRoutes++; continue; }

    // ── Block 1: Prev period (21-31) ──
    if (b1TripCol >= 0) {
      var b1Trip  = parseFloat(row[b1TripCol]) || 0;
      var b1Price = b1PriceCol >= 0 ? parseFloat(row[b1PriceCol]) || 0 : 0;
      var b1Ath   = b1AthCol   >= 0 ? parseFloat(row[b1AthCol])   || 0 : 0;

      // บันทึกเฉพาะ row ที่มี SHIFT จริงๆ (DAY/NIGHT/EXTRA) และมีข้อมูล
      if (shift && (b1Trip > 0 || b1Ath > 0)) {
        records.push({
          company: company, sheet_name: sheetName,
          route: lastRoute, shift: shift, products: lastProd,
          period_month: prevPer,
          total_trip: b1Trip, price: b1Price, total_ath: b1Ath,
        });
      }
    }

    // ── Block 2: Curr period (1-20) ──
    if (b2TripCol >= 0) {
      var b2Trip  = parseFloat(row[b2TripCol]) || 0;
      var b2Price = b2PriceCol >= 0 ? parseFloat(row[b2PriceCol]) || 0 : 0;
      var b2Ath   = b2AthCol   >= 0 ? parseFloat(row[b2AthCol])   || 0 : 0;

      if (shift && (b2Trip > 0 || b2Ath > 0)) {
        records.push({
          company: company, sheet_name: sheetName,
          route: lastRoute, shift: shift, products: lastProd,
          period_month: currPeriod,
          total_trip: b2Trip, price: b2Price, total_ath: b2Ath,
        });
      }
    }
  }

  if (nullRoutes > 0) { issues.push(nullRoutes + " แถว Route/Shift ว่าง"); }
  issues.push("✓ " + records.length + " records | prev:"+prevPer+" curr:"+currPeriod);
  issues.push("Cols → TRIP:["+tripCols+"] ATH:["+athCols+"] PRICE:["+priceCols+"]");

  var status = nullRoutes > 0 ? "warning" : "ok";
  return {
    records: records,
    validation: { sheet: sheetName, status: status, emoji: status==="ok"?"✅":"⚠️", issues: issues },
  };
}

// ── DETECT HELPERS ────────────────────────────────────────────
function detectCompany(filename) {
  var u = filename.toUpperCase();
  if (u.indexOf("ATH") >= 0) return "ATH";
  if (u.indexOf("MSP") >= 0) return "MSP";
  return "UNKNOWN";
}

function detectPeriodMonth(filename) {
  var m = filename.match(/(\d{4})[-_]?(\d{2})/);
  if (m) return m[1] + "-" + m[2];
  var now = new Date();
  return now.getFullYear() + "-" + String(now.getMonth()+1).padStart(2,"0");
}

// ============================================================
// SECTION 4: DATA STORAGE
// ============================================================

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(1,1,1,headers.length).setBackground("#1a1a2e").setFontColor("#f0b429").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveRecords(records) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var headers = ["company","sheet_name","route","shift","products","period_month","total_trip","price","total_ath","uploaded_at"];
  var sheet   = getOrCreateSheet(ss, CONFIG.DATA_SHEET, headers);
  var now     = new Date().toISOString();
  var rows    = records.map(function(r){
    return [r.company,r.sheet_name,r.route,r.shift,r.products,r.period_month,r.total_trip,r.price,r.total_ath,now];
  });
  if (rows.length > 0)
    sheet.getRange(sheet.getLastRow()+1,1,rows.length,headers.length).setValues(rows);
}

function logUpload(filename, company, periodMonth, validations) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var headers = ["filename","company","period_month","status","sheets_ok","sheets_warn","sheets_error","uploaded_at"];
  var sheet   = getOrCreateSheet(ss, CONFIG.LOG_SHEET, headers);
  var ok   = validations.filter(function(v){return v.status==="ok";}).length;
  var warn = validations.filter(function(v){return v.status==="warning";}).length;
  var err  = validations.filter(function(v){return v.status==="error";}).length;
  sheet.appendRow([filename,company,periodMonth,err>0?"error":warn>0?"warning":"ok",ok,warn,err,new Date().toISOString()]);
}

// ============================================================
// SECTION 5: KPI AGGREGATION + MoM/YoY
// ============================================================

function _buildAggMap(rows, headers) {
  var aggMap = {};
  rows.forEach(function(row){
    var obj={}; headers.forEach(function(h,i){obj[h]=row[i];});
    var key=obj.company+"|"+obj.route+"|"+obj.period_month;
    if (!aggMap[key]) aggMap[key]={company:obj.company,route:obj.route,period_month:obj.period_month,total_trips:0,total_revenue:0,prices:[]};
    aggMap[key].total_trips   += parseFloat(obj.total_trip)||0;
    aggMap[key].total_revenue += parseFloat(obj.total_ath) ||0;
    if (parseFloat(obj.price)>0) aggMap[key].prices.push(parseFloat(obj.price));
  });
  return aggMap;
}

function _addMomYoy(summaryRows, aggMap) {
  summaryRows.forEach(function(row){
    var p=row.period_month.split("-"); var y=parseInt(p[0]); var m=parseInt(p[1]);
    var pm=m===1?(y-1)+"-12":y+"-"+String(m-1).padStart(2,"0");
    var py=(y-1)+"-"+String(m).padStart(2,"0");
    var prev=aggMap[row.company+"|"+row.route+"|"+pm];
    var pyear=aggMap[row.company+"|"+row.route+"|"+py];
    function pct(a,b){return b>0?((a-b)/b*100).toFixed(2):"";}
    row.mom_revenue_pct=pct(row.total_revenue,prev ?prev.total_revenue :0);
    row.mom_trips_pct  =pct(row.total_trips,  prev ?prev.total_trips  :0);
    row.yoy_revenue_pct=pct(row.total_revenue,pyear?pyear.total_revenue:0);
    row.yoy_trips_pct  =pct(row.total_trips,  pyear?pyear.total_trips  :0);
  });
}

function recalculateAll()    { _doRecalculate(true);  }
function recalculateSilent() { _doRecalculate(false); }

function _doRecalculate(showAlert) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var ds=ss.getSheetByName(CONFIG.DATA_SHEET);
  if (!ds||ds.getLastRow()<2){if(showAlert)SpreadsheetApp.getUi().alert("ยังไม่มีข้อมูล");return;}
  var data=ds.getDataRange().getValues();
  var headers=data[0].map(function(h){return String(h);});
  var aggMap=_buildAggMap(data.slice(1),headers);
  var summaryRows=Object.values(aggMap).map(function(agg){
    return Object.assign({},agg,{avg_price:agg.prices.length>0?agg.prices.reduce(function(a,b){return a+b;},0)/agg.prices.length:0});
  });
  _addMomYoy(summaryRows,aggMap);

  var sh=["company","route","period_month","total_trips","total_revenue","avg_price","mom_revenue_pct","mom_trips_pct","yoy_revenue_pct","yoy_trips_pct","computed_at"];
  var ss2=ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if(ss2) ss2.clearContents();
  ss2=getOrCreateSheet(ss,CONFIG.SUMMARY_SHEET,sh);
  var now=new Date().toISOString();
  var out=summaryRows.map(function(r){
    return [r.company,r.route,r.period_month,Number(r.total_trips.toFixed(0)),Number(r.total_revenue.toFixed(2)),Number(r.avg_price.toFixed(2)),r.mom_revenue_pct,r.mom_trips_pct,r.yoy_revenue_pct,r.yoy_trips_pct,now];
  });
  if(out.length>0){
    ss2.getRange(2,1,out.length,sh.length).setValues(out);
    ss2.getRange(2,4,out.length,3).setNumberFormat("#,##0.00");
    ss2.getRange(2,7,out.length,4).setNumberFormat('+0.00;-0.00;"–"');
  }
  SpreadsheetApp.flush();
  if(showAlert) SpreadsheetApp.getUi().alert("✅ คำนวณ KPI เสร็จแล้ว\n"+out.length+" แถวถูกอัปเดต");
}

// ============================================================
// SECTION 6: DATA QUALITY CHECK
// ============================================================

function runQualityCheck() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var ds=ss.getSheetByName(CONFIG.DATA_SHEET);
  if(!ds){SpreadsheetApp.getUi().alert("❌ ไม่พบ Sheet KPI_DATA");return;}
  var data=ds.getDataRange().getValues(); var h=data[0]; var rows=data.slice(1);
  var ri=h.indexOf("route"),pi=h.indexOf("price"),ai=h.indexOf("total_ath");
  var issues=[];
  var nr=rows.filter(function(r){return !r[ri];}).length;
  var zp=rows.filter(function(r){return parseFloat(r[pi])<=0;}).length;
  var za=rows.filter(function(r){return parseFloat(r[ai])<=0;}).length;
  if(nr>0) issues.push("⚠️ "+nr+" แถว Route ว่าง");
  if(zp>0) issues.push("⚠️ "+zp+" แถว Price = 0");
  if(za>0) issues.push("⚠️ "+za+" แถว Revenue = 0");
  SpreadsheetApp.getUi().alert("🔍 Data Quality\n\n"+(issues.length?issues.join("\n"):"✅ "+rows.length+" แถว ผ่านทั้งหมด"));
}

// ============================================================
// SECTION 7: API สำหรับ Dashboard HTML
// ============================================================

function getSummaryData(company, yearMonth) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if(!sheet||sheet.getLastRow()<2) return {error:"ยังไม่มีข้อมูล",routes:[],totals:{}};
  var data=sheet.getDataRange().getValues();
  var headers=data[0].map(function(h){return String(h);});
  var rows=data.slice(1).map(function(row){var obj={};headers.forEach(function(h,i){obj[h]=row[i];});return obj;});
  if(company&&company!=="all") rows=rows.filter(function(r){return r.company===company;});
  if(yearMonth) rows=rows.filter(function(r){return String(r.period_month)===yearMonth;});
  var rev=rows.reduce(function(s,r){return s+(parseFloat(r.total_revenue)||0);},0);
  var trp=rows.reduce(function(s,r){return s+(parseFloat(r.total_trips)  ||0);},0);
  var prices=rows.map(function(r){return parseFloat(r.avg_price);}).filter(function(p){return p>0;});
  var avgP=prices.length>0?prices.reduce(function(a,b){return a+b;},0)/prices.length:0;
  return {totals:{revenue:rev.toFixed(2),trips:trp,avgPrice:avgP.toFixed(2)},routes:rows};
}

function getCompareData(company, periodMonth) {
  var cur=getSummaryData(company,periodMonth);
  var p=periodMonth.split("-"); var y=parseInt(p[0]); var m=parseInt(p[1]);
  var prevM=m===1?(y-1)+"-12":y+"-"+String(m-1).padStart(2,"0");
  var prevY=(y-1)+"-"+String(m).padStart(2,"0");
  var mom=getSummaryData(company,prevM); var yoy=getSummaryData(company,prevY);
  function pct(a,b){return b>0?((a-b)/b*100).toFixed(1):null;}
  var cR=parseFloat(cur.totals&&cur.totals.revenue||0);
  var cT=parseFloat(cur.totals&&cur.totals.trips||0);
  return {
    current:cur,mom:{period:prevM,data:mom},yoy:{period:prevY,data:yoy},
    momSummary:{revenuePct:pct(cR,parseFloat(mom.totals&&mom.totals.revenue||0)),tripsPct:pct(cT,parseFloat(mom.totals&&mom.totals.trips||0))},
    yoySummary:{revenuePct:pct(cR,parseFloat(yoy.totals&&yoy.totals.revenue||0)),tripsPct:pct(cT,parseFloat(yoy.totals&&yoy.totals.trips||0))},
  };
}

function getTrendData(company) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if(!sheet||sheet.getLastRow()<2) return [];
  var data=sheet.getDataRange().getValues();
  var headers=data[0].map(function(h){return String(h);});
  var rows=data.slice(1).map(function(row){var obj={};headers.forEach(function(h,i){obj[h]=row[i];});return obj;});
  if(company&&company!=="all") rows=rows.filter(function(r){return r.company===company;});
  var bm={};
  rows.forEach(function(r){var pm=String(r.period_month);if(!bm[pm]) bm[pm]={period_month:pm,revenue:0,trips:0};bm[pm].revenue+=parseFloat(r.total_revenue)||0;bm[pm].trips+=parseFloat(r.total_trips)||0;});
  return Object.values(bm).sort(function(a,b){return a.period_month.localeCompare(b.period_month);}).slice(-12);
}

function getAvailablePeriods() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if(!sheet||sheet.getLastRow()<2) return [];
  var data=sheet.getDataRange().getValues();
  var pmIdx=data[0].map(function(h){return String(h);}).indexOf("period_month");
  if(pmIdx<0) return [];
  var seen={},periods=[];
  data.slice(1).forEach(function(r){var v=String(r[pmIdx]);if(v&&!seen[v]){seen[v]=true;periods.push(v);}});
  return periods.sort().reverse();
}

function getQualityStatus() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(CONFIG.LOG_SHEET);
  if(!sheet||sheet.getLastRow()<2) return [];
  var data=sheet.getDataRange().getValues();
  var headers=data[0].map(function(h){return String(h);});
  return data.slice(1).map(function(row){var obj={};headers.forEach(function(h,i){obj[h]=row[i];});return obj;}).slice(-10).reverse();
}

// ============================================================
// SECTION 8: EXPORT
// ============================================================

function exportToExcel() {
  var id=SpreadsheetApp.getActiveSpreadsheet().getId();
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput("<script>window.open('https://docs.google.com/spreadsheets/d/"+id+"/export?format=xlsx');google.script.host.close();<\/script>"),
    "กำลังดาวน์โหลด..."
  );
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
