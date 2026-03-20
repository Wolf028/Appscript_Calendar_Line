function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ลงบันทึกราชการ - LINE LIFF')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://your-icon-link.png');
}


function sendLineGroupMessage(groupId, message) {
  var channelToken = '/B6D4gzQxwqrOXAR/lp0AN+Y74ut/tcZ6dxruD09dqwarJxARfMW1zJ2fkvLpIf7i+9aY4bGn6OXanEhu8k4DWi7T/AdU1pqZ6bUueqm4mqmq38OYwr+AOLgMOkyLE0pR6VaQslYGurI0bRIMlVZFgdB04t89/1O/w1cDnyilFU='; // ใส่ Channel Access Token ของคุณ
  var url = 'https://api.line.me/v2/bot/message/push';
  var payload = {
    "to": groupId,
    "messages": [{
      "type": "text",
      "text": message
    }]
  };
  var options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + channelToken
    },
    "payload": JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}

// --- ส่วนตรวจสอบการเข้าสู่ระบบ (Login) ---
function checkLogin(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users'); // ดึงข้อมูลจาก Sheet 'Users' ที่มีอยู่แล้ว
  
  // 1. (Option) ตั้งรหัส Admin สูงสุดไว้ในโค้ดกันเหนียว (เผื่อใน Sheet ไม่มี)
  if (username === 'admin' && password === '1234') {
    return 'ADMIN';
  }

  // 2. วนลูปเช็คชื่อใน Sheet Users
  var data = sheet.getDataRange().getValues(); // ดึงข้อมูลมาทั้งหมด
  // สมมติว่า Column A = Username, Column B = Password
  for (var i = 1; i < data.length; i++) { // เริ่มวนลูปบรรทัดที่ 2 (ข้ามหัวตาราง)
    if (data[i][0] == username && data[i][1] == password) {
      return 'USER'; // เจอข้อมูลตรงกัน ส่งค่ากลับว่าเป็น USER
    }
  }

  return 'FAIL'; // ไม่เจอชื่อ
}

// --- ส่วนลงทะเบียนสมาชิกใหม่ (Register) ---
function registerUser(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  
  // 1. เช็คก่อนว่าชื่อซ้ำไหม?
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == username) {
      return 'DUPLICATE'; // ชื่อซ้ำ แจ้งกลับไป
    }
  }

  // 2. ถ้าไม่ซ้ำ ให้บันทึกลง Sheet (ต่อท้ายแถวล่าสุด)
  sheet.appendRow([username, password]);
  return 'SUCCESS';
}

// ฟังก์ชันเดียวจัดการทั้ง บันทึกใหม่ และ แก้ไข
function processForm(data) { 
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Data');
    var start = new Date(data.startDate);
    var end = new Date(data.endDate);

    // ตั้งเวลาในคอลัมน์วันเริ่ม/วันสิ้นสุดให้ไม่มีเวลา
    start.setHours(0, 0, 0, 0);
    end.setHours(0, 0, 0, 0);

    if (data.rowId && data.rowId !== "") {
      var rowIndex = parseInt(data.rowId);
      if (isNaN(rowIndex) || rowIndex < 2) return "Error: ไม่พบแถวข้อมูล";
      var rng = sheet.getRange(rowIndex, 2, 1, 6);

      rng.setValues([[data.fullname, data.position, start, end, data.location, data.recorder]]);
      sheet.getRange(rowIndex, 4, 1, 2).setNumberFormat('dd/MM/yyyy'); // ฟอร์แมตเฉพาะคอลัมน์วันที่

    } else {
      // ✅ Timestamp เดิม (แสดงวัน-เวลา)
      var newRow = [
        new Date(), // Timestamp แบบเต็ม (มีเวลา)
        data.fullname,
        data.position,
        start,
        end,
        data.location,
        data.recorder
      ];

      sheet.appendRow(newRow);

      // เซ็ตฟอร์แมตเฉพาะคอลัมน์วันที่ให้เป็น dd/MM/yyyy
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 4, 1, 2).setNumberFormat('dd/MM/yyyy'); // Start & End date
    }

    var msg = "บันทึกออกราชการใหม่\n" +
              "ชื่อ👤: " + data.fullname + "\n" +
              "ตำแหน่ง📌: " + data.position + "\n" +
              "สถานที่📍: " + data.location + "\n" +
              "วันที่📅: " + data.startDate + " ถึง " + data.endDate;
    var groupId = 'C10cab2e4f02f44e4558db5d85b2d1f78'; // ใส่ User ID ของผู้รับ
    sendLineGroupMessage(groupId, msg);

    return "success";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

function deleteData(rowNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data');
  sheet.deleteRow(rowNumber);
  return "success";
}

// --- ฟังก์ชันดึงข้อมูล (แก้ใหม่ให้ชัวร์เรื่องวันที่) ---
function getData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  if (!sheet) return []; // ถ้าหา Sheet ไม่เจอ ให้ส่งค่าว่างกลับไป กัน Error

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // ถ้ามีแค่หัวตาราง ไม่มีข้อมูล ให้ส่งค่าว่าง

  data.shift(); // ลบหัวตารางออก
  
  var result = [];
  
  // วนลูปเช็คข้อมูลทีละแถว
  data.forEach(function(row, index) {
    // เช็คว่าต้องมี "ชื่อ" และ "วันที่เริ่ม" ถึงจะเอามาแสดง (ป้องกันแถวว่างๆ)
    if(row[1] && row[3]) {
      
      // แปลงวันที่เป็น String มาตรฐาน YYYY-MM-DD (แก้ปัญหา Timezone เพี้ยน)
      var startStr = "";
      var endStr = "";
      try {
         startStr = Utilities.formatDate(new Date(row[3]), "GMT+7", "yyyy-MM-dd");
         endStr = Utilities.formatDate(new Date(row[4]), "GMT+7", "yyyy-MM-dd");
      } catch(e) {
         startStr = ""; endStr = "";
      }

      result.push({
        rowId: index + 2, // เลขบรรทัดใน Sheet
        name: row[1],
        position: row[2],
        start: startStr,  // ส่งเป็น Text YYYY-MM-DD
        end: endStr,      // ส่งเป็น Text YYYY-MM-DD
        location: row[5],
        // จัด Format สวยๆ ไว้แสดงในตาราง
        startDisplay: formatDateThai(row[3]), 
        endDisplay: formatDateThai(row[4])
      });
    }
  });
  
  return result;
}

// ฟังก์ชันช่วยแปลงวันที่เป็นไทย
function formatDateThai(date) {
  if (!date) return "-";
  return Utilities.formatDate(new Date(date), "GMT+7", "dd/MM/yyyy");
}

function testSendToGroup() {
  var groupId = 'C10cab2e4f02f44e4558db5d85b2d1f78'; // ใส่ Group ID ที่ได้มา
  var msg = 'แจ้งเตือนจากบอท: ทดสอบส่งข้อความเข้าไลน์กลุ่มสำเร็จ!';
  sendLineGroupMessage(groupId, msg);
}
