function doPost(e) {
  const sheet = SpreadsheetApp.openById("16HkfnOE-c0xzhkkOlQCYDL1DaDjzARouNTNd5eDiBdM")
                .getSheetByName("Data");
  const data = sheet.getDataRange().getValues();
  
  const body = JSON.parse(e.postData.contents);
  const replyToken = body.events[0].replyToken;
  const userMessage = body.events[0].message.text.trim();
  
  let replyText = "";
  let results = [];

  // ======================
  // ถ้าพิมพ์ขึ้นต้นด้วย "วันที่ "
  // ======================
  if (userMessage.startsWith("วันที่ ")) {

    const searchDate = userMessage.replace("วันที่ ", "").trim();

    for (let i = 1; i < data.length; i++) {
      if (formatDate(data[i][3]) == searchDate) {

        results.push(
          "👤 " + data[i][1] +
          "\n📌 " + data[i][2] +
          "\n📅 เริ่ม: " + formatDate(data[i][3]) +
          "\n📅 สิ้นสุด: " + formatDate(data[i][4]) +
          "\n📍 " + data[i][5]
        );
      }
    }

    if (results.length > 0) {
      replyText = "📅 ข้อมูลวันที่ " + searchDate + "\n\n" +
                  results.join("\n\n----------------\n\n");
    } else {
      replyText = "❌ ไม่พบข้อมูลวันที่ " + searchDate;
    }
  }

  // ======================
  // กรณีพิมพ์ชื่ออย่างเดียว
  // ======================
  else {

    const searchName = userMessage.toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const sheetName = String(data[i][1]).toLowerCase();

      if (sheetName.includes(searchName)) {

        results.push(
          "👤 " + data[i][1] +
          "\n📌 " + data[i][2] +
          "\n📅 เริ่ม: " + formatDate(data[i][3]) +
          "\n📅 สิ้นสุด: " + formatDate(data[i][4]) +
          "\n📍 " + data[i][5]
        );
      }
    }

    if (results.length > 0) {
      replyText = "🔎 ผลการค้นหา \"" + userMessage + "\"\n\n" +
                  results.join("\n\n----------------\n\n");
    } else {
      replyText = "❌ ไม่พบข้อมูลที่มีคำว่า \"" + userMessage + "\"";
    }
  }

  replyLine(replyToken, replyText);
}

// แปลงวันที่
function formatDate(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, "Asia/Bangkok", "dd/MM/yyyy");
  }
  return date;
}

function replyLine(replyToken, message) {
  const url = "https://api.line.me/v2/bot/message/reply";
  const token = "/B6D4gzQxwqrOXAR/lp0AN+Y74ut/tcZ6dxruD09dqwarJxARfMW1zJ2fkvLpIf7i+9aY4bGn6OXanEhu8k4DWi7T/AdU1pqZ6bUueqm4mqmq38OYwr+AOLgMOkyLE0pR6VaQslYGurI0bRIMlVZFgdB04t89/1O/w1cDnyilFU=";

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + token
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{
        type: "text",
        text: message
      }]
    })
  };

  UrlFetchApp.fetch(url, options);
}