function onOpen() {
  var subMenus = [{name:"Update",functionName:"update"}];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Weather", subMenus);
}

/**
 * Updates spreadsheet with current weather
 */
function update() {
  // locations by e.g. ZIP or city,country
  var locations = ["10001","basel, switzerland","singapore"];
  var values = [];
  values[0] = new Date();
  var i = 1;
  for (var l = 0; l < 4; l++) {
    var locData = getCurrentWeather(locations[l]);
    for (var d = 0; d < 6; d++) {
      values[i] = locData[d];
      i++;
    }
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow(values);  
}

/**
 * Gets the current weather for a location
 */
function getCurrentWeather(location) {
  var url = "http://www.google.com/ig/api?weather=" + location;
  var xml = UrlFetchApp.fetch(url).getContentText();
  var doc = Xml.parse(xml, true);
  var current = doc.xml_api_reply.weather.current_conditions
  var condition = current.condition.getAttribute("data").getValue();
  var tempC = current.temp_c.getAttribute("data").getValue();
  var tempF = current.temp_f.getAttribute("data").getValue();
  var humidity = current.humidity.getAttribute("data").getValue().substring(10);
  var icon = current.icon.getAttribute("data").getValue().substring(19);
  var wind = current.wind_condition.getAttribute("data").getValue().substring(5);
  var result = [condition,tempF,tempC,humidity,icon,wind];
  return result;
}

