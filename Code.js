function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}

function crawlToSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const url =
    "https://www.dtro.or.kr/open_content_new/ko/OpenApi/culture.php?ST_DT=20200822&ED_DT=20200922"; // 크롤링할 정적 페이지 주소

  const response = UrlFetchApp.fetch(url);
  const xml = response.getContentText("UTF-8");
  const document = XmlService.parse(xml);
  const apiDataItems = findElementsByName(document.getRootElement(), "apiData");
  const fields = [
    "ST_DATE",
    "ED_DATE",
    "ST_TIME",
    "ED_TIME",
    "STD_NM",
    "EVENT_NM",
    "EVENT_CNTNT",
    "EVENT_HOST",
  ];

  if (apiDataItems.length === 0) {
    throw new Error("응답 XML에서 apiData 항목을 찾지 못했습니다.");
  }

  const rows = apiDataItems.map(function (item) {
    return fields.map(function (field) {
      return item.getChildText(field) || "";
    });
  });

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(fields);
  }

  const firstRow = sheet.getLastRow() + 1;
  sheet.getRange(firstRow, 1, rows.length, fields.length).setValues(rows);
  Logger.log("%s개 항목을 시트에 저장했습니다.", rows.length);
}

function findElementsByName(element, name) {
  let matches = [];

  if (element.getName() === name) {
    matches.push(element);
  }

  element.getChildren().forEach(function (child) {
    matches = matches.concat(findElementsByName(child, name));
  });

  return matches;
}
