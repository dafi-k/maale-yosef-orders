// ════════════════════════════════════════════════════════════
// Google Apps Script — שליחת הזמנות ריהוט | מעלה יוסף
// גרסה 2 — שולח קובץ Excel כקובץ מצורף
// ════════════════════════════════════════════════════════════
// הוראות פריסה:
// 1. פתחו: https://script.google.com
// 2. לחצו "פרויקט חדש" → שנו שם לפרויקט (למשל: "Maale Yosef Orders")
// 3. מחקו הכל → הדביקו את הקוד הזה כולו
// 4. לחצו "שמור" (Ctrl+S)
// 5. לחצו "פרוס" ← "פריסה חדשה"
//    - סוג פריסה: "יישום אינטרנט"
//    - הפעל כ: "אני" (Me)
//    - מי יכול לגשת: "כל אחד" (Anyone)
// 6. לחצו "פרוס" → אשרו הרשאות אם נדרש
// 7. העתיקו את כתובת ה-Web App URL
// 8. ב-index.html, החליפו את הטקסט PLACEHOLDER_APPS_SCRIPT_URL
//    בכתובת שהעתקתם
// ════════════════════════════════════════════════════════════

// Parse POST body manually to bypass e.parameter 8190-char limit (needed for xlsxBase64)
function parsePostBody(e) {
  var p = {};
  if (e && e.parameter) {
    for (var k in e.parameter) { p[k] = e.parameter[k]; }
  }
  if (e && e.postData && e.postData.contents) {
    e.postData.contents.split('&').forEach(function(pair) {
      var idx = pair.indexOf('=');
      if (idx > 0) {
        try {
          var key = decodeURIComponent(pair.substring(0, idx).replace(/\+/g, ' '));
          var val = decodeURIComponent(pair.substring(idx + 1).replace(/\+/g, ' '));
          p[key] = val;
        } catch(ex) {}
      }
    });
  }
  return p;
}

function doPost(e) {
  try {
    var params = parsePostBody(e);

    var to       = params.to           || '';
    var cc       = params.cc           || '';
    var subject  = params.subject      || 'הזמנת ריהוט — מעלה יוסף';
    var body     = params.body         || '';
    var cabin    = params.cabin        || '';
    var contact  = params.contact      || '';
    var phone    = params.phone        || '';
    var total    = params.total        || '';
    var date     = params.date         || new Date().toLocaleDateString('he-IL');
    var xlsxB64  = params.xlsxBase64   || '';
    var xlsxName = params.xlsxFilename || ('הזמנה_' + cabin + '.xlsx');

    if (!to) {
      Logger.log('doPost: missing "to" field');
      return ContentService.createTextOutput('error:missing-to');
    }

    var htmlBody = buildHtmlEmail(cabin, contact, phone, to, body, total, date);

    var mailOptions = {
      to:       to,
      cc:       cc,
      subject:  subject,
      body:     body,
      htmlBody: htmlBody
    };

    // Attach the Excel file if base64 data was provided
    if (xlsxB64) {
      try {
        var xlsxBytes = Utilities.base64Decode(xlsxB64);
        var xlsxBlob  = Utilities.newBlob(
          xlsxBytes,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          xlsxName
        );
        mailOptions.attachments = [xlsxBlob];
      } catch (attachErr) {
        Logger.log('Attachment creation failed: ' + attachErr.toString());
        // Continue without attachment rather than failing entirely
      }
    }

    MailApp.sendEmail(mailOptions);
    Logger.log('Email sent to: ' + to + (cc ? ' | CC: ' + cc : ''));

    return ContentService.createTextOutput('OK');

  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return ContentService.createTextOutput('error:' + err.toString());
  }
}

// ────────────────────────────────────────────
// בונה גוף מייל HTML מעוצב
// ────────────────────────────────────────────
function buildHtmlEmail(cabin, contact, phone, email, plainBody, total, date) {
  // Parse item lines from plain text body
  var lines     = plainBody.split('\n');
  var tableRows = '';
  var inItems   = false;
  var totalLine = '';
  var balanceLine = '';

  lines.forEach(function(line) {
    if (line.indexOf('──') !== -1) { inItems = !inItems; return; }
    if (line.indexOf('סה"כ הזמנה:') !== -1)  { totalLine   = line; return; }
    if (line.indexOf('יתרה לתשלום') !== -1)   { balanceLine = line; return; }
    if (inItems && line.trim()) {
      var parts = line.split('|').map(function(s){ return s.trim(); });
      var total = parts.length;
      tableRows += '<tr>' + parts.map(function(p, i) {
        var isNumeric = (i >= total - 3); // last 3 cols: כמות, מחיר, סה"כ
        var style = 'padding:7px 10px;border-bottom:1px solid #e8d5b7;font-size:0.88rem;' +
                    (isNumeric ? 'text-align:center;direction:ltr;' : '') +
                    (i === total - 1 ? 'font-weight:600;' : '');
        return '<td style="' + style + '">' + p + '</td>';
      }).join('') + '</tr>';
    }
  });

  var balanceHtml = balanceLine
    ? '<p style="color:#c0392b;font-weight:700;font-size:1rem;margin:8px 0 0;">' + balanceLine + '</p>'
    : '';

  return [
    '<div dir="rtl" style="font-family:Arial,sans-serif;max-width:720px;margin:0 auto;background:#f5f0e8;padding:24px;border-radius:12px;">',

    // Header
    '<div style="background:linear-gradient(135deg,#6B4E2A,#8B6840);color:#fff;padding:24px;border-radius:8px;text-align:center;margin-bottom:20px;">',
    '<h2 style="margin:0 0 6px;font-size:1.4rem;">הזמנת ריהוט — רענון צימרים 2025</h2>',
    '<p style="margin:0;opacity:0.85;font-size:0.9rem;">מועצה אזורית מעלה יוסף — גליל מערבי</p>',
    '</div>',

    // Identity box
    '<div style="background:#fff;border-radius:8px;padding:16px 20px;margin-bottom:16px;border:1px solid #ddd3c0;">',
    '<table style="width:100%;border-collapse:collapse;">',
    '<tr><td style="padding:4px 0;color:#7a6e5f;font-size:0.83rem;width:120px;">שם הצימר:</td><td style="font-weight:700;font-size:1rem;">' + cabin + '</td></tr>',
    '<tr><td style="padding:4px 0;color:#7a6e5f;font-size:0.83rem;">איש קשר:</td><td>' + contact + '</td></tr>',
    '<tr><td style="padding:4px 0;color:#7a6e5f;font-size:0.83rem;">נייד:</td><td>' + phone + '</td></tr>',
    '<tr><td style="padding:4px 0;color:#7a6e5f;font-size:0.83rem;">אי-מייל:</td><td>' + email + '</td></tr>',
    '<tr><td style="padding:4px 0;color:#7a6e5f;font-size:0.83rem;">תאריך:</td><td>' + date + '</td></tr>',
    '</table></div>',

    // Items table
    '<div style="background:#fff;border-radius:8px;overflow:hidden;border:1px solid #ddd3c0;margin-bottom:16px;">',
    '<table style="width:100%;border-collapse:collapse;">',
    '<thead><tr style="background:#6B4E2A;color:#fff;font-size:0.82rem;">',
    '<th style="padding:9px 10px;text-align:right;">פריט</th>',
    '<th style="padding:9px 10px;text-align:right;">מק"ט</th>',
    '<th style="padding:9px 10px;text-align:right;">תיאור</th>',
    '<th style="padding:9px 10px;text-align:right;">כמות</th>',
    '<th style="padding:9px 10px;text-align:right;">מחיר</th>',
    '<th style="padding:9px 10px;text-align:right;">סה"כ</th>',
    '</tr></thead>',
    '<tbody>',
    tableRows || '<tr><td colspan="6" style="padding:14px;text-align:center;color:#999;">אין פריטים</td></tr>',
    '</tbody></table></div>',

    // Totals
    '<div style="background:#fff;border-radius:8px;padding:14px 20px;border:1px solid #ddd3c0;">',
    total ? '<p style="font-size:1.1rem;font-weight:700;color:#6B4E2A;margin:0 0 4px;">סה"כ הזמנה: ₪' + Number(total).toLocaleString('he-IL') + '</p>' : '',
    '<p style="color:#7a6e5f;font-size:0.88rem;margin:4px 0 0;">תקציב מועצה: ₪9,500</p>',
    balanceHtml,
    '</div>',

    // Footer note
    '<p style="color:#aaa;font-size:0.75rem;margin-top:18px;text-align:center;">',
    'מייל זה נשלח אוטומטית ממערכת קטלוג ריהוט מועצה אזורית מעלה יוסף.<br>',
    'קובץ Excel של ההזמנה מצורף למייל זה.',
    '</p>',
    '</div>'
  ].join('');
}

// ────────────────────────────────────────────
// פונקציית בדיקה — הרץ ידנית מהעורך כדי לבדוק
// ────────────────────────────────────────────
function testDoPost() {
  var fakeE = {
    parameter: {
      to:           'dafioz@gmail.com',
      cc:           'lior@contrast-il.com,marketing@contrast-il.com',
      subject:      'בדיקה — הזמנת ריהוט | צימר בדיקה',
      cabin:        'צימר בדיקה',
      contact:      'ישראל ישראלי',
      phone:        '050-1234567',
      email:        'dafioz@gmail.com',
      total:        '7500',
      date:         new Date().toLocaleDateString('he-IL'),
      xlsxBase64:   '',
      xlsxFilename: 'הזמנה_בדיקה.xlsx',
      body: [
        'סיכום הזמנת ריהוט — צימר בדיקה',
        'תאריך: ' + new Date().toLocaleDateString('he-IL'),
        'איש קשר: ישראל ישראלי | נייד: 050-1234567',
        '',
        'פרטי ההזמנה:',
        '─────────────────────────────────────────',
        'A1 | כורסא מסתובבת | כמות: 2 | ₪2,599 | ₪5,198',
        'C1 | כיסא דמוי עור | כמות: 4 | ₪799 | ₪3,196',
        '─────────────────────────────────────────',
        'סה"כ הזמנה: ₪8,394'
      ].join('\n')
    }
  };
  var result = doPost(fakeE);
  Logger.log('Result: ' + result.getContent());
}
