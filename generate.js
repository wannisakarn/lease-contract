// สร้างสัญญาเช่าอาคาร (.docx) จากไฟล์ข้อมูล JSON
// ใช้งาน: node generate.js [data.json] [output.docx]
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell,
  WidthType, BorderStyle, PageNumber, Footer } = require('docx');

const dataFile = process.argv[2] || 'data.example.json';
const d = JSON.parse(fs.readFileSync(dataFile, 'utf8'));
const out = process.argv[3] || path.join('output', `สัญญาเช่าอาคาร_${(d.property.address_no || 'template').replace(/\//g, '-')}.docx`);

// ---------- helpers ----------
const B = '..........................';
const v = (x, blank = B) => (x === undefined || x === null || x === '' ? blank : String(x));
const num = n => (n || n === 0) ? Number(n).toLocaleString('en-US') : B;
const num2 = n => (n || n === 0) ? Number(n).toLocaleString('en-US', { minimumFractionDigits: 2 }) : B;
const TH = ['๐','๑','๒','๓','๔','๕','๖','๗','๘','๙'];
const thNum = n => String(n).replace(/\d/g, c => TH[c]);

function bahtText(n) {
  if (!n && n !== 0) return B;
  const digits = ['ศูนย์','หนึ่ง','สอง','สาม','สี่','ห้า','หก','เจ็ด','แปด','เก้า'];
  const pos = ['','สิบ','ร้อย','พัน','หมื่น','แสน'];
  const read = s => {
    let out = '';
    const len = s.length;
    for (let i = 0; i < len; i++) {
      const dg = +s[i], p = len - i - 1;
      if (dg === 0) continue;
      if (p === 1 && dg === 1) out += 'สิบ';
      else if (p === 1 && dg === 2) out += 'ยี่สิบ';
      else if (p === 0 && dg === 1 && len > 1) out += 'เอ็ด';
      else out += digits[dg] + pos[p];
    }
    return out;
  };
  const conv = s => {
    if (s.length > 6) return conv(s.slice(0, -6)) + 'ล้าน' + read(s.slice(-6));
    return read(s);
  };
  const [i, f = '00'] = Number(n).toFixed(2).split('.');
  const baht = +i === 0 ? 'ศูนย์' : conv(i);
  return baht + 'บาท' + (+f === 0 ? 'ถ้วน' : conv(f.replace(/^0/, '')) + 'สตางค์');
}

const F = d.font || 'TH Sarabun New', SZ = (d.font_size || 16) * 2;
const r = (t, o = {}) => new TextRun({ text: t, font: F, size: SZ, ...o });
const b = t => r(t, { bold: true });
const P = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [r(runs)],
  alignment: AlignmentType.THAI_DISTRIBUTE, spacing: { after: 60, line: 300 }, ...o });
const IND = runs => P(Array.isArray(runs) ? runs : [r(runs)], { indent: { firstLine: 720 } });
let clauseNo = 0;
const C = runs => { clauseNo++; return P([b(`ข้อ ${thNum(clauseNo)} `), ...(Array.isArray(runs) ? runs : [r(runs)])], { indent: { firstLine: 720 } }); };
const center = (t, o = {}) => new Paragraph({ children: [r(t, o)], alignment: AlignmentType.CENTER, spacing: { after: 60 } });
const right = t => new Paragraph({ children: [r(t)], alignment: AlignmentType.RIGHT, spacing: { after: 0 } });
const sign = (role, name) => [
  new Paragraph({ children: [r(`ลงชื่อ......................................................${role}`)], alignment: AlignmentType.RIGHT, spacing: { before: 360, after: 0 } }),
  new Paragraph({ children: [r(`(${v(name, '......................................................')})`)], alignment: AlignmentType.RIGHT, indent: { right: 700 }, spacing: { after: 0 } })];

// ---------- data ----------
const L = d.landlord, T = d.tenant, PR = d.property, RT = d.rent, DP = d.deposit, BK = d.bank;
const propAddr = `${v(PR.address_no)} ${v(PR.address_detail, '')}`.trim();
const rent = RT.monthly;
const adv = RT.advance_months || 0;

const kids = [
  center('สัญญาเช่าอาคาร', { bold: true, size: 44 }),
  new Paragraph({ spacing: { after: 120 } }),
  right(`ทำที่ ${v(d.signed_at_line1)}`),
  ...(d.signed_at_line2 ? [right(d.signed_at_line2)] : []),
  right(`วันที่ ${v(d.contract_date)}`),
  new Paragraph({ spacing: { after: 120 } }),
  IND([r('สัญญาเช่าฉบับนี้ทำขึ้นระหว่าง '), b(v(L.name)), r(` อยู่บ้านเลขที่ ${v(L.address)} ถือบัตรประจำตัวประชาชนเลขที่ ${v(L.id_no)} ซึ่งต่อไปในสัญญาจะเรียกว่า `), b('“ผู้ให้เช่า”'),
    r(' ฝ่ายหนึ่ง กับ '), b(v(T.name)), r(` อยู่บ้านเลขที่ ${v(T.address)} ถือบัตรประจำตัวประชาชนเลขที่ ${v(T.id_no)} ออกให้ ณ ${v(T.id_issued_at)} เมื่อวันที่ ${v(T.id_issued_date)} ซึ่งต่อไปในสัญญาจะเรียกว่า `), b('“ผู้เช่า”'),
    r(' อีกฝ่ายหนึ่ง คู่สัญญาทั้งสองฝ่ายตกลงทำสัญญากันโดยมีเงื่อนไขและรายละเอียดดังต่อไปนี้')]),
  C(`ผู้ให้เช่าตกลงให้เช่าและผู้เช่าตกลงเช่าอาคาร คือ อาคารเลขที่ ${propAddr}${PR.house_code ? ` (เลขรหัสประจำบ้าน ${PR.house_code})` : ''}${PR.building_type ? ` ลักษณะอาคารเป็น${PR.building_type}` : ''} เพื่อเป็นประโยชน์ในการ${v(PR.purpose)}`),
  C([r(`ระยะเวลาการเช่า มีกำหนด ${v(d.term.text)} เริ่มตั้งแต่วันที่ `), b(v(d.term.start)), r(' สิ้นสุดวันที่ '), b(v(d.term.end)),
    r(` หากผู้เช่าประสงค์จะต่ออายุสัญญาเช่า ผู้เช่าจะต้องแจ้งให้ผู้ให้เช่าทราบเป็นลายลักษณ์อักษรล่วงหน้าไม่น้อยกว่า ${v(d.term.renewal_notice)} ก่อนวันครบกำหนดสัญญา และการต่อสัญญาให้กระทำเป็นรายปี ทุกปี ตามเงื่อนไขที่คู่สัญญาตกลงกัน`)]),
  C([r(`ผู้เช่าตกลงชำระเงินค่าเช่าให้ผู้ให้เช่าเป็นรายเดือน โดยกำหนดชำระล่วงหน้าภายในวันที่ ${v(RT.due_day, '.......')} ของแต่ละเดือน ทุกเดือน ในอัตราค่าเช่าเดือนละ `), b(`${num(rent)} บาท (${bahtText(rent)})`),
    ...(adv ? [r(` ทั้งนี้ ในวันทำสัญญานี้ ผู้เช่าได้ชำระค่าเช่าล่วงหน้า ${adv} เดือน เป็นเงิน ${num(rent * adv)} บาท (${bahtText(rent * adv)})${RT.advance_period ? ` สำหรับงวด${RT.advance_period}` : ''} ให้แก่ผู้ให้เช่าแล้ว`)] : [])]),
  C([r(`การชำระค่าเช่านั้น ผู้เช่าจะต้องโอนเงินค่าเช่าเข้าบัญชี${v(BK.account_type, 'เงินฝาก')}ของผู้ให้เช่า ${v(BK.bank)} สาขา${v(BK.branch)} ชื่อบัญชี ${v(BK.account_name)} เลขที่บัญชี `), b(v(BK.account_no)), r(' และส่งหลักฐานการโอนเงินให้ผู้ให้เช่าทราบทุกครั้ง')]),
  C('ผู้เช่าจะใช้ทรัพย์สินที่เช่าเพื่อการอย่างอื่นนอกจากที่ระบุไว้ในข้อ ๑ แห่งสัญญานี้ได้ ต่อเมื่อได้รับความยินยอมเป็นหนังสือจากผู้ให้เช่า'),
  C('ผู้เช่าจะดูแลรักษาอาคารที่เช่าให้อยู่ในสภาพที่ดี สะอาด เรียบร้อย อย่างเช่นวิญญูชนจะพึงปฏิบัติในการรักษาทรัพย์สินของตน ด้วยค่าใช้จ่ายของตนเอง'),
  C('ผู้เช่าจะโอนสิทธิการเช่าตามสัญญานี้หรือจะนำทรัพย์ที่เช่าไปให้ผู้อื่นเช่าช่วงไม่ได้ เว้นแต่ผู้เช่าจะได้รับความยินยอมจากผู้ให้เช่าเป็นลายลักษณ์อักษรเสียก่อน'),
  C('ผู้เช่าจะไม่แก้ไข เปลี่ยนแปลง หรือต่อเติมอาคาร เว้นแต่จะได้รับความยินยอมจากผู้ให้เช่าเป็นลายลักษณ์อักษรเสียก่อน ทรัพย์สินที่ติดตรึงตรากับอาคารอันเป็นผลจากการตกแต่งภายใน หรือการแก้ไขเปลี่ยนแปลงใดๆ ที่ผู้เช่าได้กระทำไปโดยได้รับความยินยอมแล้วนั้น ให้ตกเป็นกรรมสิทธิ์ของผู้ให้เช่า โดยผู้เช่าจะไม่เรียกร้องค่าชดเชยใดๆ ทั้งสิ้น'),
  C('ผู้เช่าต้องรับผิดในบรรดาความเสียหายหรือบุบสลายใดๆ อันเกิดขึ้นแก่ทรัพย์สินที่เช่า เพราะความผิดของผู้เช่าหรือผู้เช่าช่วงหรือบุคคลซึ่งอยู่กับผู้เช่า'),
  C('ผู้เช่าจะยินยอมให้ผู้ให้เช่าหรือตัวแทนที่ได้รับมอบอำนาจจากผู้ให้เช่า เข้าตรวจตราอาคารได้เป็นครั้งคราวในระยะเวลาอันสมควร โดยแจ้งให้ผู้เช่าทราบล่วงหน้า'),
  C('ผู้เช่าตกลงชำระค่าไฟฟ้า ค่าน้ำประปา ค่าใช้โทรศัพท์ ค่าอินเทอร์เน็ต และค่าบริการอื่นๆ ที่เกิดขึ้นจากการใช้ประโยชน์ของอาคารที่เช่านี้ ซึ่งเรียกเก็บโดยหน่วยงานของราชการหรือผู้ให้บริการ ตลอดอายุการเช่า'),
  C('ผู้เช่าตกลงชำระค่าภาษีต่างๆ ที่เกิดขึ้นและเกี่ยวกับทรัพย์สินที่เช่าตลอดอายุการเช่า'),
  C([r(`ผู้เช่าตกลงวางเงินประกัน (เงินมัดจำ) จำนวน ${v(DP.months, '....')} เดือน ไว้ให้แก่ผู้ให้เช่าในวันทำสัญญานี้ เป็นเงิน `), b(`${num(DP.amount)} บาท (${bahtText(DP.amount)})`),
    r(`${DP.payment_ref ? ` ${DP.payment_ref}` : ''} และผู้ให้เช่าได้รับเงินประกันดังกล่าวไว้เรียบร้อยแล้ว เงินประกันดังกล่าวเป็นเงินประกันความรับผิดตามสัญญาของผู้เช่า ซึ่งผู้ให้เช่าจะคืนให้แก่ผู้เช่าในวันที่ผู้เช่าได้ย้ายบริวารและทรัพย์สินออกจากสถานที่เช่าเรียบร้อยแล้วหลังจากสิ้นสุดสัญญาเช่า และไม่ติดค้างชำระค่าเช่า ค่าสาธารณูปโภค หรือมีหนี้สินอื่นค้างชำระผู้ให้เช่า ทั้งนี้ ผู้ให้เช่ามีสิทธิหักเงินประกันเพื่อชดใช้ค่าเสียหายหรือหนี้ที่ค้างชำระได้ เงินประกันนี้ไม่อาจนำมาหักเป็นค่าเช่าเดือนใดๆ ได้`)]),
  C('การประกันทรัพย์สินของผู้เช่าในอาคารจะกระทำได้ต่อเมื่อได้รับความยินยอมหรืออนุญาตจากผู้ให้เช่าเป็นลายลักษณ์อักษรก่อน ระหว่างอายุสัญญาการเช่านี้ ผู้เช่าต้องทำประกันอัคคีภัยเพื่อคุ้มครองอาคารที่เช่ากับบริษัทที่ได้รับความยินยอมจากผู้ให้เช่า ผู้เช่าต้องเป็นผู้จ่ายเงินค่าเบี้ยประกัน โดยต้องระบุให้ผู้ให้เช่าเป็นผู้รับผลประโยชน์ และต้องให้ผู้ให้เช่าเป็นผู้เก็บกรมธรรม์'),
  C('ผู้ให้เช่าจะใช้สิทธิบอกกล่าวให้ผู้เช่าปฏิบัติตามสัญญานี้ หรือบอกเลิกสัญญาหรือเรียกค่าเสียหายกับผู้เช่าได้ เมื่อผู้เช่าผิดสัญญาข้อหนึ่งข้อใดตามสัญญานี้'),
  C('การเช่าตามสัญญานี้ย่อมสิ้นสุดลงก่อนครบกำหนดระยะเวลาที่ระบุไว้ในสัญญา เมื่อปรากฏว่าทรัพย์สินที่เช่าพินาศโดยสิ้นเชิงหรือเป็นส่วนใหญ่เพราะอัคคีภัย หรือภัยอื่นใด'),
  C(`ผู้เช่ายอมชดใช้ดอกเบี้ยในอัตราร้อยละ ${v(d.late_interest_pct, '....')} ต่อปี ของยอดเงินค่าเช่าที่ค้างชำระผู้ให้เช่า และค่าใช้จ่ายต่างๆ ที่ผู้ให้เช่าต้องเสียไปเพื่อการทวงถามให้ชำระเงินค่าเช่าอีกด้วย`),
  C('คู่สัญญาตกลงให้สัญญานี้เลิกกันเมื่อผู้เช่าถูกพิทักษ์ทรัพย์หรือล้มละลายตามกฎหมาย'),
  C(`ผู้เช่าต้องขนย้ายทรัพย์สินและบริวารของผู้เช่าและส่งมอบทรัพย์สินที่เช่าคืนให้แก่ผู้ให้เช่าในสภาพปกติทันทีเมื่อสัญญานี้สิ้นสุดลงหรือเลิกกัน หากผู้เช่าไม่ปฏิบัติตามความในข้อนี้ ผู้เช่ายอมชดใช้ค่าปรับให้แก่ผู้ให้เช่าในอัตราวันละ ${num(d.overstay_penalty_per_day)} บาท (${bahtText(d.overstay_penalty_per_day)}) นับแต่วันที่สัญญาเช่านี้เลิกกันหรือสัญญาเช่านี้สิ้นสุดลง จนกว่าจะมีการส่งมอบทรัพย์สินที่เช่าคืนให้แก่ผู้ให้เช่าแล้ว`),
  C('ในวันทำสัญญานี้ผู้เช่าได้ตรวจตราทรัพย์สินที่เช่าแล้วเห็นว่ามีสภาพปกติดี'),
  C('การส่งหนังสือบอกกล่าวใดๆ จากผู้ให้เช่าถึงผู้เช่า ให้ส่งตามที่อยู่ของผู้เช่า ดังนี้'),
  IND([b(v(d.notice_address_tenant, '................................................................................................'))]),
  IND('หากผู้ให้เช่าได้ส่งจดหมายลงทะเบียนไปยังที่อยู่ของผู้เช่าตามที่อยู่ข้างต้นนี้ ถือว่าผู้เช่าได้รับหนังสือดังกล่าว หากผู้เช่าประสงค์จะเปลี่ยนแปลงที่อยู่ ผู้เช่าต้องแจ้งให้ผู้ให้เช่าทราบถึงการเปลี่ยนแปลงนั้นๆ'),
  C('การส่งหนังสือบอกกล่าวใดๆ จากผู้เช่าถึงผู้ให้เช่า ให้ส่งตามที่อยู่ของผู้ให้เช่า ดังนี้'),
  IND([b(v(d.notice_address_landlord, '................................................................................................'))]),
  IND('หากผู้เช่าได้ส่งจดหมายลงทะเบียนไปยังที่อยู่ของผู้ให้เช่าตามที่อยู่ข้างต้นนี้ ถือว่าผู้ให้เช่าได้รับหนังสือดังกล่าว หากผู้ให้เช่าประสงค์จะเปลี่ยนแปลงที่อยู่ ผู้ให้เช่าต้องแจ้งให้ผู้เช่าทราบถึงการเปลี่ยนแปลงนั้นๆ'),
  C('การแก้ไข เพิ่มเติมหรือเปลี่ยนแปลงสัญญานี้ ไม่ว่าส่วนหนึ่งส่วนใดหรือทั้งหมด จะต้องทำเป็นหนังสือ และลงลายมือชื่อของคู่สัญญาทั้งสองฝ่ายไว้เป็นสำคัญจึงจะมีผลใช้บังคับระหว่างคู่สัญญาได้'),
  IND('สัญญานี้ทำขึ้นเป็นสองฉบับมีข้อความถูกต้องตรงกัน ผู้ให้เช่าได้ส่งมอบทรัพย์สินที่เช่าให้แก่ผู้เช่าแล้ว คู่สัญญาเห็นว่าถูกต้องตามเจตนาแห่งตน และทั้งสองฝ่ายเข้าใจข้อความแห่งสัญญานี้ทั้งหมดดีแล้ว จึงลงลายมือชื่อไว้เป็นหลักฐานสำคัญต่อหน้าพยาน และต่างฝ่ายต่างยึดถือไว้ฝ่ายละฉบับ'),
  ...sign('ผู้ให้เช่า', L.name),
  ...sign('ผู้เช่า', T.name),
  ...sign('พยาน', (d.witnesses || [])[0]),
  ...sign('พยาน', (d.witnesses || [])[1]),
  new Paragraph({ pageBreakBefore: true, children: [b('เอกสารแนบท้ายสัญญาเช่า')], alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
  ...(d.attachments || []).map((a, i) => P(`${thNum(i + 1)}. ${a}`)),
  new Paragraph({ spacing: { before: 240 }, children: [b('สรุปการชำระเงินในวันทำสัญญา')] }),
];

const W = [1000, 4700, 2600], tw = W.reduce((a, c) => a + c);
const bd = { style: BorderStyle.SINGLE, size: 4, color: '808080' }, borders = { top: bd, bottom: bd, left: bd, right: bd };
const cell = (t, i, bold) => new TableCell({ borders, width: { size: W[i], type: WidthType.DXA }, margins: { left: 100, right: 100 },
  children: [new Paragraph({ alignment: i === 2 ? AlignmentType.RIGHT : i === 0 ? AlignmentType.CENTER : AlignmentType.LEFT, children: [r(t, { bold })] })] });
const row = (a, bold) => new TableRow({ children: a.map((t, i) => cell(t, i, bold)) });
const rows = [row(['ลำดับ', 'รายการ', 'จำนวนเงิน (บาท)'], true)];
let total = 0, k = 1;
if (DP.amount) { rows.push(row([String(k++), `เงินประกัน (มัดจำ) ${v(DP.months, '')} เดือน`, num2(DP.amount)])); total += DP.amount; }
if (adv && rent) { rows.push(row([String(k++), `ค่าเช่าล่วงหน้า ${adv} เดือน`, num2(rent * adv)])); total += rent * adv; }
rows.push(row(['', `รวม (${total ? bahtText(total) : B})`, total ? num2(total) : B], true));
kids.push(new Table({ width: { size: tw, type: WidthType.DXA }, columnWidths: W, rows }));
kids.push(...sign('ผู้ให้เช่า', L.name), ...sign('ผู้เช่า', T.name));

const doc = new Document({ styles: { default: { document: { run: { font: F, size: SZ } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1440, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [r('หน้า ', { size: 24 }),
      new TextRun({ children: [PageNumber.CURRENT], font: F, size: 24 }), r(' / ', { size: 24 }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: F, size: 24 })] })] }) },
    children: kids }] });

Packer.toBuffer(doc).then(buf => { fs.mkdirSync(path.dirname(out), { recursive: true }); fs.writeFileSync(out, buf); console.log('สร้างไฟล์แล้ว:', out); });
