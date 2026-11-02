/**
 * سكربت Google Apps Script — استقبال حركات حميدا وكتابتها في شيت "رصيد مع حميدا".
 *
 * طريقة التنصيب (مرة واحدة فقط):
 * 1) افتح الشيت الأساسي في المتصفح.
 * 2) من القائمة: الإضافات ← Apps Script  (Extensions ← Apps Script).
 * 3) احذف أي كود موجود في الملف الذي يفتح، وألصق هذا الملف كامله مكانه.
 * 4) غيّر قيمة SHEET_NAME أدناه إذا كان اسم التبويب مختلفًا، وغيّر SECRET إلى كلمة مرور من اختيارك.
 * 5) من زر "نشر" (Deploy) ← "نشر جديد" (New deployment) ← اختر النوع "تطبيق ويب" (Web app).
 *    - "من يستطيع الوصول" (Who has access): اختر "أي شخص" (Anyone).
 *    - اضغط نشر (Deploy)، وامنح الصلاحيات المطلوبة لحسابك عند السؤال.
 * 6) انسخ الرابط (Web app URL) الذي يظهر، وضعه في ملف index.html للتطبيق، في المتغيّر
 *    APPS_SCRIPT_URL، وضع كلمة المرور نفسها في SYNC_SECRET.
 */

var SHEET_NAME = 'رصيد مع حميدا';
var SECRET = 'GHAYYER_HETHI'; // غيّرها لكلمة مرور من اختيارك، واستخدم نفسها في التطبيق

// يحوّل نص التاريخ القادم من التطبيق (يوم/شهر/سنة) إلى تاريخ حقيقي صريح (سنة، شهر، يوم)،
// بدل ما نكتب النص زي ما هو ونخلي جوجل شيت يحاول يخمّن صيغته بنفسه. التخمين هذا هو سبب
// مشكلة ظهرت فعليًا: لما يكون اليوم والشهر كلاهما ≤12 (مثلاً 01/11/2026 ليوم 1 نوفمبر)،
// جوجل شيت بيفسّرها أحيانًا كـ"شهر/يوم" (حسب لغة/منطقة الشيت) ويحوّلها لتاريخ مختلف تمامًا
// (11 يناير بدل 1 نوفمبر). ببناء كائن Date صريح هنا، ما بضل أي مجال للتخمين أو الالتباس.
function parseAppDateString(str){
  var parts = String(str||'').trim().split('/');
  if(parts.length !== 3) return null;
  var day = Number(parts[0]), month = Number(parts[1]), year = Number(parts[2]);
  if(!day || !month || !year) return null;
  var d = new Date(year, month - 1, day);
  return isNaN(d.getTime()) ? null : d;
}

function doPost(e) {
  var result = { ok: false };
  try {
    var body = JSON.parse(e.postData.contents);

    if (!body.secret || body.secret !== SECRET) {
      result.error = 'unauthorized';
      return respond(result);
    }

    if (body.action === 'delete') {
      var delRow = Number(body.row);
      if (!delRow || delRow < 5) {
        result.error = 'invalid_row';
        return respond(result);
      }
      var ssDel = SpreadsheetApp.getActiveSpreadsheet();
      var sheetDel = ssDel.getSheetByName(SHEET_NAME);
      if (!sheetDel) {
        result.error = 'sheet_not_found';
        return respond(result);
      }
      // نفرّغ محتوى الصف بدل حذفه بالكامل، عشان ما تنزاح أرقام باقي الصفوف (وبالتالي باقي الحركات المتزامنة)
      sheetDel.getRange(delRow, 2, 1, 4).clearContent();
      result.ok = true;
      return respond(result);
    }

    if (body.action === 'update') {
      var updRow = Number(body.row);
      if (!updRow || updRow < 5) {
        result.error = 'invalid_row';
        return respond(result);
      }
      var updAmount = Number(body.amount);
      var updType = body.type === 'in' ? 'in' : 'out';
      var updDesc = String(body.desc || '').trim();
      var updDate = String(body.date || '').trim();
      if (!updAmount || updAmount <= 0 || !updDesc || !updDate) {
        result.error = 'invalid_data';
        return respond(result);
      }
      var ssUpd = SpreadsheetApp.getActiveSpreadsheet();
      var sheetUpd = ssUpd.getSheetByName(SHEET_NAME);
      if (!sheetUpd) {
        result.error = 'sheet_not_found';
        return respond(result);
      }
      var updDateObj = parseAppDateString(updDate);
      sheetUpd.getRange(updRow, 2, 1, 4).setValues([[
        updType === 'in' ? updAmount : '',
        updType === 'out' ? updAmount : '',
        updDesc,
        updDateObj || updDate
      ]]);
      sheetUpd.getRange(updRow, 5).setNumberFormat('dd/mm/yyyy'); // العمود E (التاريخ)
      result.ok = true;
      return respond(result);
    }

    var amount = Number(body.amount);
    var type = body.type === 'in' ? 'in' : 'out';
    var desc = String(body.desc || '').trim();
    var date = String(body.date || '').trim();

    if (!amount || amount <= 0 || !desc || !date) {
      result.error = 'invalid_data';
      return respond(result);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      result.error = 'sheet_not_found';
      return respond(result);
    }

    // ترتيب الأعمدة في الشيت: B=وارد, C=ناقص, D=البيان, E=التاريخ
    var lastRow = sheet.getLastRow();
    var newRow = lastRow + 1;
    var dateObj = parseAppDateString(date);
    sheet.getRange(newRow, 2, 1, 4).setValues([[
      type === 'in' ? amount : '',
      type === 'out' ? amount : '',
      desc,
      dateObj || date
    ]]);
    sheet.getRange(newRow, 5).setNumberFormat('dd/mm/yyyy'); // العمود E (التاريخ): يمنع أي تفسير ملتبس لاحقًا

    result.ok = true;
    result.row = newRow;
    return respond(result);
  } catch (err) {
    result.error = String(err);
    return respond(result);
  }
}

// يُستخدم لسحب كل الحركات المسجّلة في الشيت، ليستطيع أي جهاز (حاسوب أو جوال) عرض نفس
// البيانات الموحّدة، لا فقط الجهاز اللي أضافها. البيانات تبدأ من الصف 5 (بعد صفوف الترويسة).
function doGet(e) {
  var result = { ok: false };
  try {
    var secret = e.parameter.secret;
    if (!secret || secret !== SECRET) {
      result.error = 'unauthorized';
      return respond(result);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      result.error = 'sheet_not_found';
      return respond(result);
    }

    var lastRow = sheet.getLastRow();
    var entries = [];
    if (lastRow >= 5) {
      var data = sheet.getRange(5, 2, lastRow - 4, 4).getValues(); // B:E من الصف 5
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var inAmt = row[0], outAmt = row[1], desc = row[2], dateVal = row[3];
        if (!inAmt && !outAmt && !desc) continue; // صف فاضي
        // بعض خلايا التاريخ بالشيت محفوظة كتاريخ حقيقي لا كنص، فتُنسَّق هنا بنفس صيغة التطبيق: يوم/شهر/سنة
        var dateStr = (dateVal instanceof Date)
          ? Utilities.formatDate(dateVal, ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy')
          : String(dateVal || '');
        entries.push({
          row: 5 + i,
          type: inAmt ? 'in' : 'out',
          amount: Number(inAmt || outAmt || 0),
          desc: String(desc || ''),
          date: dateStr
        });
      }
    }

    result.ok = true;
    result.entries = entries;
    return respond(result);
  } catch (err) {
    result.error = String(err);
    return respond(result);
  }
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
