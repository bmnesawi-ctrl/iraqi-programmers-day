// Google Apps Script — يستقبل توثيق الحضور ويضيفه كصف جديد في Google Sheet
const SHEET_NAME = 'المشاركون';
const HEADERS = ['التسلسل', 'رقم التوثيق', 'الاسم الرباعي', 'رقم الهاتف', 'المحافظة', 'الدولة', 'الانتماء للنقابة', 'الانطباع عن الاحتفالية', 'تاريخ التوثيق'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.setRightToLeft(true);
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#18286A').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('D:D').setNumberFormat('@'); // الهاتف كنص
    sh.getRange('I:I').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('H:H').setWrap(true);
    sh.setColumnWidths(1, HEADERS.length, 140);
    sh.setColumnWidth(3, 260);
    sh.setColumnWidth(8, 380);
  }
  return sh;
}

// حماية بسيطة من أن يُفسَّر النص كمعادلة داخل Excel
function safe_(v, max) {
  v = String(v == null ? '' : v).trim().slice(0, max || 120);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = getSheet_();
    const phone = String(d.phone || '').trim();

    // منع تسجيل الهاتف نفسه مرتين
    const last = sh.getLastRow();
    if (last > 1) {
      const phones = sh.getRange(2, 4, last - 1, 1).getValues().flat().map(String);
      if (phones.indexOf(phone) !== -1) {
        return out_({ ok: false, message: 'تم توثيق حضور بهذا الرقم مسبقًا. يرجى التواصل مع إدارة الفعالية إذا كنت تعتقد أن ذلك غير صحيح.' });
      }
    }

    const seq = last; // الصف الأول للعناوين
    const reg = 'IPD-' + String(seq).padStart(6, '0');
    sh.appendRow([seq, reg, safe_(d.name), phone, safe_(d.governorate), safe_(d.country), safe_(d.union), safe_(d.impression, 500), new Date()]);
    sh.getRange(sh.getLastRow(), 4).setNumberFormat('@').setValue(phone);
    return out_({ ok: true, registration_number: reg });
  } catch (err) {
    return out_({ ok: false, message: 'حدث خطأ أثناء الحفظ.' });
  } finally {
    lock.releaseLock();
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
