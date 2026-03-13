// noinspection GrazieStyle

const fs   = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const folderPath = 'C:\\Users\\asus\\Acwa\\QIPP - QIPP Mail Ingest Temp';
const outputFile = path.join(__dirname, 'plant_data.json');
const CAPACITY   = 3883.2;

const MONTHS  = ["january","february","march","april","may","june","july",
    "august","september","october","november","december"];
const SHORT_M = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];

// Only accept Excel date serials in the valid 2025-2027 range (~45900-46500)
function excelSerialToDate(val) {
    if (typeof val !== 'number') return null;
    if (val < 45900 || val > 46500) return null;
    try {
        const d = XLSX.SSF.parse_date_code(val);
        return `${String(d.d).padStart(2,'0')}.${String(d.m).padStart(2,'0')}.${d.y}`;
    } catch { return null; }
}

function dateFromFilename(fileName) {
    const m1 = fileName.match(/(\d{2})[.-]([a-zA-Z]+)[.-](\d{4})/);
    if (m1) {
        let mi = MONTHS.indexOf(m1[2].toLowerCase());
        if (mi === -1) mi = SHORT_M.indexOf(m1[2].toLowerCase().substring(0, 3));
        if (mi >= 0) return `${m1[1].padStart(2,'0')}.${String(mi+1).padStart(2,'0')}.${m1[3]}`;
    }
    const m2 = fileName.match(/(\d{2})[.-](\d{2})[.-](\d{4})/);
    if (m2) return `${m2[1].padStart(2,'0')}.${m2[2].padStart(2,'0')}.${m2[3]}`;
    return null;
}

// Unit definitions
const UNIT_MAP = [
    {group:'G1', unit:'11', type:'GT', row:19},
    {group:'G1', unit:'12', type:'GT', row:20},
    {group:'G1', unit:'10', type:'ST', row:21},
    {group:'G2', unit:'21', type:'GT', row:22},
    {group:'G2', unit:'23', type:'GT', row:23},
    {group:'G2', unit:'20', type:'ST', row:24},
    {group:'G3', unit:'31', type:'GT', row:25},
    {group:'G3', unit:'32', type:'GT', row:26},
    {group:'G3', unit:'30', type:'ST', row:27},
    {group:'G4', unit:'41', type:'GT', row:28},
    {group:'G4', unit:'42', type:'GT', row:29},
    {group:'G4', unit:'40', type:'ST', row:30},
    {group:'G5', unit:'51', type:'GT', row:31},
    {group:'G5', unit:'52', type:'GT', row:32},
    {group:'G5', unit:'50', type:'ST', row:33},
    {group:'G6', unit:'61', type:'GT', row:34},
    {group:'G6', unit:'62', type:'GT', row:35},
    {group:'G6', unit:'60', type:'ST', row:36},
];

const allData = {};

function getOrCreate(dateStr) {
    if (!allData[dateStr]) {
        allData[dateStr] = {
            Date:        dateStr,
            Generation:  null,
            NetGen:      null,
            Load:        null,
            PLF:         null,
            Efficiency:  null,
            HeatRate:    null,
            Fuel:        null,
            Aux:         null,
            MFEQH:       null,
            Emissions:   { NOx: null, SOx: null, CO: null },
            Water:       { ROProduction: null },
            AirIntakeDP: null,
            Units:       []
        };
    }
    return allData[dateStr];
}

function sv(v) {
    const n = Number(v);
    return (!isNaN(n) && isFinite(n) && n !== 0) ? n : null;
}

try {
    const files = fs.readdirSync(folderPath)
        .filter(f => f.endsWith('.xlsx') && !f.startsWith('~$'));
    console.log(`🔍 Scanning ${files.length} Excel files...`);

    const subset = files.filter(f => {
        const u = f.toUpperCase();
        return ['DAILY ACTUAL ENERGY PRODUCED', 'TIMERS-COUNTERS',
            'AIR INTAKE', 'ENVIRONMENT REPORT', 'DAILY OPERATION'].some(k => u.includes(k));
    });
    console.log(`⏳ Processing ${subset.length} relevant files...`);

    let cnt = 0;
    subset.forEach(fileName => {
        cnt++;
        if (cnt % 10 === 0) console.log(`⏳ Processed ${cnt}/${subset.length} files...`);

        const nameUpper = fileName.toUpperCase();
        try {
            const wb = XLSX.readFile(path.join(folderPath, fileName), {
                sheetRows: 60, cellFormula: false, cellHTML: false,
                cellText: false, cellNF: false
            });

            if (nameUpper.includes('DAILY ACTUAL ENERGY PRODUCED')) {
                const rows = XLSX.utils.sheet_to_json(
                    wb.Sheets[wb.SheetNames[0]], { header: 1 });

                let dateStr = null;
                (rows[3] || []).forEach(c => { if (!dateStr) dateStr = excelSerialToDate(c); });
                if (!dateStr) dateStr = dateFromFilename(fileName);
                if (!dateStr) return;

                const day = getOrCreate(dateStr);

                let totalGen = 0, totalLoad = 0, hrs = 0;
                for (let i = 6; i < 30; i++) {
                    const r   = rows[i] || [];
                    const gen = Number(r[7]);
                    const ld  = Number(r[6]);
                    if (!isNaN(gen) && gen > 0) { totalGen += gen; totalLoad += ld; hrs++; }
                }
                if (hrs > 0 && totalGen > 0 && !day.Generation) {
                    day.Generation = totalGen;
                    day.Load       = totalLoad / hrs;
                }
            }

            else if (nameUpper.includes('TIMERS-COUNTERS')) {
                const dateStr = dateFromFilename(fileName);
                if (!dateStr) return;
                const day   = getOrCreate(dateStr);
                const sheet = wb.Sheets['COMPILED'] || wb.Sheets[wb.SheetNames[0]];
                const rows  = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                let mfeqh = 0;
                rows.forEach(r => {
                    if (!r || r[0] == null) return;
                    const v = Number(r[3]);
                    if (!isNaN(v) && v > 0 && v <= 24) mfeqh += v;
                });
                if (mfeqh > 0) day.MFEQH = parseFloat(mfeqh.toFixed(1));
            }

            else if (nameUpper.includes('AIR INTAKE')) {
                const dateStr = dateFromFilename(fileName);
                if (!dateStr) return;
                const day  = getOrCreate(dateStr);
                const rows = XLSX.utils.sheet_to_json(
                    wb.Sheets[wb.SheetNames[0]], { header: 1 });
                let sum = 0, n = 0;
                rows.forEach(r => {
                    if (r && typeof r[0] === 'string' && r[0].toUpperCase().includes('GT')) {
                        const dp = Number(r[4]);
                        if (!isNaN(dp) && dp > 0 && dp < 50) { sum += dp; n++; }
                    }
                });
                if (n > 0) day.AirIntakeDP = parseFloat((sum / n).toFixed(3));
            }

            else if (nameUpper.includes('ENVIRONMENT REPORT')) {
                const dateStr = dateFromFilename(fileName);
                if (!dateStr) return;
                const day  = getOrCreate(dateStr);
                const rows = XLSX.utils.sheet_to_json(
                    wb.Sheets[wb.SheetNames[0]], { header: 1 });
                rows.forEach(r => {
                    if (!r || typeof r[0] !== 'string') return;
                    const p   = r[0].trim().toUpperCase();
                    const v6  = Number(r[6])  || 0;
                    const v11 = Number(r[11]) || 0;
                    if (p === 'NOX' && (v6 || v11)) day.Emissions.NOx = parseFloat(((v6 + v11) / 2).toFixed(3));
                    if (p === 'SOX' && (v6 || v11)) day.Emissions.SOx = parseFloat(((v6 + v11) / 2).toFixed(3));
                    if (p === 'CO'  && (v6 || v11)) day.Emissions.CO  = parseFloat(((v6 + v11) / 2).toFixed(3));
                });
            }

            else if (nameUpper.includes('DAILY OPERATION')) {
                const dateStr = dateFromFilename(fileName);
                if (!dateStr) return;
                const day = getOrCreate(dateStr);

                const sheetName = wb.SheetNames.find(
                    n => n.toLowerCase().includes('daily operation')) || wb.SheetNames[0];
                const rows = XLSX.utils.sheet_to_json(
                    wb.Sheets[sheetName], { header: 1 });

                const r6  = rows[6]  || [];
                const r13 = rows[13] || [];
                const r15 = rows[15] || [];
                const r40 = rows[40] || [];
                const r42 = rows[42] || [];

                if (sv(r6[4]))  day.Load       = sv(r6[4]);
                if (sv(r13[3])) day.Generation = sv(r13[3]);
                if (sv(r13[4])) day.PLF        = parseFloat(Number(r13[4]).toFixed(2));
                if (sv(r13[8])) day.Aux         = sv(r13[8]);
                if (sv(r15[4])) day.Fuel        = sv(r15[4]);

                // GasHV and NetGen math completely stripped here

                if (!day.PLF && day.Load) day.PLF = parseFloat((day.Load / CAPACITY * 100).toFixed(2));

                const ro = Number(r40[8]);
                if (!isNaN(ro) && ro > 0) day.Water.ROProduction = ro;

                day.TempMax  = sv(r42[3]);
                day.TempMin  = sv(r42[5]);
                day.TempAvg  = sv(r42[6]);
                day.MaxRH    = sv(r42[8]);
                day.MinRH    = sv(r42[10]);
                day.WindSpeed= sv(r42[12]);

                day.Units = [];
                UNIT_MAP.forEach(u => {
                    const r = rows[u.row] || [];
                    const gen  = sv(r[3]);
                    const load = sv(r[2]);
                    const mf   = sv(r[4]);
                    day.Units.push({
                        Group:      u.group,
                        Unit:       u.unit,
                        Type:       u.type,
                        Load:       load,
                        Generation: gen,
                        MFEQH:      mf
                    });
                });
            }

        } catch (e) {
            console.warn(`⚠️  ${fileName}: ${e.message}`);
        }
    });

    const finalData = Object.values(allData)
        .filter(d => d.Date && d.Generation != null && d.Generation > 0)
        .sort((a, b) => {
            const t = s => {
                const p = s.split('.');
                return new Date(`${p[2]}-${p[1]}-${p[0]}`).getTime();
            };
            return t(a.Date) - t(b.Date);
        });

    fs.writeFileSync(outputFile, JSON.stringify(finalData, null, 2));
    console.log(`\n✅ Success — ${finalData.length} days written to ${outputFile}`);

} catch (e) {
    console.error('❌ Fatal Error:', e);
}

async function init() {
    try {
        // Fetching your specific JSON file
        const resp = await fetch('plant_data.json');
        if (!resp.ok) throw new Error('HTTP ' + resp.status);
        const rawData = await resp.json();

        const dataArray = Array.isArray(rawData) ? rawData : (rawData.records || rawData.data || []);

        ALL = dataArray.map((r) => {
            // Flatten nested datatree (if present) so KPIs are visible
            const flatData = r.datatree ? {...r, ...r.datatree} : r;

            // Parse Date
            const d = flatData.Date || flatData.date || flatData.Day;
            const dp = d ? String(d).trim() : null;
            r.d = dp ? (() => {
                const parts = dp.split(/[\/\-.]/).map(p => p.trim());
                if (parts.length === 3 && parts[0].length <= 2) {
                    return new Date(+parts[2], +parts[1] - 1, +parts[0]); // dd.MM.yyyy
                }
                return new Date(dp);
            })() : null;

            const keys = Object.keys(flatData);

            // Robust Key Detection for Heat Rate
            const hrKey = keys.find((k) => {
                const lk = k.toLowerCase();
                return lk.includes('heat rate') || lk.includes('heatrate') || /\bhr\b/.test(lk) || (lk.includes('kj') && lk.includes('kwh'));
            });

            // Robust Key Detection for Efficiency
            const effKey = keys.find((k) => {
                const lk = k.toLowerCase();
                return lk.includes('efficiency') || lk.includes('eff');
            });

            // Robust Key Detection for MFEQH
            const mfKey = keys.find((k) => k.toLowerCase().includes('mfeqh'));

            // Map KPI fields onto r so the existing UI code sees them
            r.HeatRate = hrKey ? cleanNum(flatData[hrKey]) : null;
            r.Efficiency = effKey ? cleanNum(flatData[effKey]) : null;
            r.MFEQH = mfKey ? cleanNum(flatData[mfKey]) : null;

            // Normalize main KPIs if they live under datatree
            if (flatData.Generation !== undefined) r.Generation = cleanNum(flatData.Generation);
            if (flatData.Load !== undefined) r.Load = cleanNum(flatData.Load);
            if (flatData.PLF !== undefined) r.PLF = cleanNum(flatData.PLF);
            if (flatData.Fuel !== undefined) r.Fuel = cleanNum(flatData.Fuel);
            if (flatData['Fuel Gas'] !== undefined) r.Fuel = cleanNum(flatData['Fuel Gas']);

            return r;
        }).filter((r) => r.d instanceof Date && !isNaN(r.d.getTime())).sort((a, b) => a.d - b.d);

        VIEW = [...ALL];

        // Trigger your existing UI build functions
        if (typeof buildTgl === 'function') buildTgl();
        if (typeof sync === 'function') sync();
        if (typeof draw === 'function') draw();

    } catch (e) {
        console.error('Data load error:', e);
        const el = document.getElementById('dash-info');
        if (el) el.textContent = 'Data not loaded - check plant_data.json';
    }
}