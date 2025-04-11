var WEATHER_API_KEY = "6b53ccde6a6579b4d492f10482074298";
var CHAMBLEE_LAT = 33.8922;
var CHAMBLEE_LON = -84.2988;

function doGet() {
  Logger.log("doGet: Starting app initialization");
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = spreadsheet.getSheetByName('Settings');
    
    if (!settingsSheet) {
      Logger.log("doGet: Settings sheet not found, creating...");
      ensureSettingsSheet();
    }
    
    // Check if validation is disabled
    const disableValidation = settingsSheet.getRange("B3").getValue();
    const validationDisabled = disableValidation.toString().toLowerCase() === "yes";
    Logger.log(`doGet: Validation disabled: ${validationDisabled}`);

    // If validation is not disabled, perform user check
    if (!validationDisabled) {
      const allowedUsers = settingsSheet.getRange('D1:D10')
        .getValues()
        .flat()
        .filter(Boolean);
      Logger.log(`doGet: Allowed users: ${allowedUsers}`);

      if (!userEmail || !allowedUsers.includes(userEmail)) {
        Logger.log(`doGet: Access denied for user: ${userEmail || 'anonymous'}`);
        return HtmlService.createHtmlOutput(`
          <h2>Access Denied</h2>
          <p>Sorry, you (${userEmail || 'anonymous'}) are not authorized to access this application.</p>
          <p>Please contact the administrator to request access.</p>
        `);
      }
    }

    Logger.log("doGet: User authorized, serving Index.html");
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Live with gusto! Inventory & Prep - Chamblee')
      .setWidth(800)
      .setHeight(600);
  } catch (error) {
    Logger.log(`doGet: Error - ${error.message}`);
    return HtmlService.createHtmlOutput(`<h2>Error</h2><p>Failed to initialize app: ${error.message}</p>`);
  }
}

function updateProductMixOverrides(updates) {
  var variableDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("D: Variable Data");
  updates.forEach(update => {
    const value = update.value !== '' ? (parseFloat(update.value) / 100) : '';
    variableDataSheet.getRange("F6:F13").getCell(update.row, 1).setValue(value);
    Logger.log(`Updated product mix override at row ${update.row} with value ${value}`);
  });
}

function getInventoryData() {
  try {
    Logger.log("getInventoryData: Starting data fetch");
    var data = {};
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("InventoryPrep");
    Logger.log("getInventoryData: Accessed InventoryPrep sheet");

    data.cookies = {
      names: sheet.getRange("B4:B6").getValues().flat(),
      amounts: sheet.getRange("C4:C6").getValues().flat()
    };
    Logger.log("getInventoryData: Fetched cookies data");

    data.medleys = {
      names: sheet.getRange("B9:B16").getValues().flat(),
      inventory: sheet.getRange("C9:E16").getValues(),
      catering: sheet.getRange("G9:K16").getValues()
    };
    Logger.log("getInventoryData: Fetched medleys data");

    data.sauces = {
      names: sheet.getRange("B19:B27").getValues().flat(),
      inventory: sheet.getRange("C19:E27").getValues(),
      quarts: sheet.getRange("G19:G27").getValues().flat()
    };
    Logger.log("getInventoryData: Fetched sauces data");

    data.garnishes = {
      names: sheet.getRange("B30:B34").getValues().flat(),
      inventory: sheet.getRange("C30:E34").getValues(),
      quarts: sheet.getRange("G30:G34").getValues().flat()
    };
    Logger.log("getInventoryData: Fetched garnishes data");

    data.vegetables = {
      names: sheet.getRange("B36:B55").getValues().flat(),
      amounts: sheet.getRange("C36:C55").getValues().flat()
    };
    Logger.log("getInventoryData: Fetched vegetables data");

    data.dressings = {
      names: sheet.getRange("B57:B64").getValues().flat(),
      amounts: sheet.getRange("C57:C64").getValues().flat()
    };
    Logger.log("getInventoryData: Fetched dressings data");

    data.cateringMedleys = {
      names: sheet.getRange("B9:B16").getValues().flat().map(name => name.trim()),
      portions: sheet.getRange("G9:K16").getValues()
    };
    Logger.log("getInventoryData: Fetched catering medleys data");

    data.calendarEvents = getCalendarEvents();
    Logger.log(`getInventoryData: Calendar events summary - Today: ${data.calendarEvents.todayEvents.length}, Upcoming: ${data.calendarEvents.upcomingEvents.length}, Ongoing: ${data.calendarEvents.ongoingEvents.length}`);
    
    var today = new Date();
    data.overview = {
      date: Utilities.formatDate(today, "EST", "MMMM dd, yyyy"),
      dayOfWeek: Utilities.formatDate(today, "EST", "EEEE"),
      calendarEvents: data.calendarEvents
    };
    Logger.log("getInventoryData: Set overview date and calendar events");

    var cacheSheet = spreadsheet.getSheetByName("WeatherCache");
    if (!cacheSheet) {
      cacheSheet = spreadsheet.insertSheet("WeatherCache");
      cacheSheet.getRange("A1").setValue("Last Updated");
      cacheSheet.getRange("B1").setValue("Temperature");
      cacheSheet.getRange("C1").setValue("Description");
      cacheSheet.getRange("D1").setValue("Hourly Data Date");
      cacheSheet.getRange("E1").setValue("Hourly Data");
    }
    Logger.log("getInventoryData: Accessed WeatherCache sheet");

    var lastUpdated = cacheSheet.getRange("A2").getValue();
    var cachedTemp = cacheSheet.getRange("B2").getValue();
    var cachedDesc = cacheSheet.getRange("C2").getValue();
    var cachedHourlyDate = cacheSheet.getRange("D2").getValue();
    var cachedHourlyData = cacheSheet.getRange("E2").getValue();
    var now = new Date();
    var oneHour = 60 * 60 * 1000;

    if (lastUpdated && (now - new Date(lastUpdated)) < oneHour && cachedTemp && cachedDesc) {
      data.overview.weather = {
        temp: Math.round(cachedTemp),
        description: cachedDesc
      };
      Logger.log("getInventoryData: Used cached weather data");
    } else {
      var weatherUrl = `http://api.openweathermap.org/data/2.5/weather?lat=${CHAMBLEE_LAT}&lon=${CHAMBLEE_LON}&appid=${WEATHER_API_KEY}&units=imperial`;
      try {
        var response = UrlFetchApp.fetch(weatherUrl);
        var json = JSON.parse(response.getContentText());
        data.overview.weather = {
          temp: Math.round(json.main.temp),
          description: json.weather[0].description
        };
        cacheSheet.getRange("A2").setValue(now.toISOString());
        cacheSheet.getRange("B2").setValue(data.overview.weather.temp);
        cacheSheet.getRange("C2").setValue(data.overview.weather.description);
        Logger.log("getInventoryData: Fetched and cached new weather data");
      } catch (e) {
        Logger.log(`getInventoryData: Weather fetch error - ${e.message}`);
        data.overview.weather = { temp: "N/A", description: "Unable to fetch weather data" };
      }
    }

    var hourlyWeatherUrl = `https://api.openweathermap.org/data/3.0/onecall?lat=${CHAMBLEE_LAT}&lon=${CHAMBLEE_LON}&exclude=current,minutely,daily,alerts&units=imperial&appid=${WEATHER_API_KEY}`;
    try {
      var hourlyResponse = UrlFetchApp.fetch(hourlyWeatherUrl);
      var hourlyJson = JSON.parse(hourlyResponse.getContentText());
      var freshHourlyData = hourlyJson.hourly.map(hour => ({
        dt: hour.dt,
        temp: hour.temp,
        weather: hour.weather[0].description
      }));
      Logger.log("getInventoryData: Fetched fresh hourly weather data");

      // Cache the hourly data for the current day
      const todayStr = Utilities.formatDate(today, "EST", "yyyy-MM-dd");
      cacheSheet.getRange("D2").setValue(todayStr);
      cacheSheet.getRange("E2").setValue(JSON.stringify(freshHourlyData));
      Logger.log("getInventoryData: Cached hourly weather data for today");
    } catch (e) {
      Logger.log(`getInventoryData: Hourly weather fetch error - ${e.message}`);
      freshHourlyData = [];
    }

    // Combine cached and fresh hourly data
    let combinedHourlyData = [];
    try {
      const cachedDateStr = cachedHourlyDate ? Utilities.formatDate(new Date(cachedHourlyDate), "EST", "yyyy-MM-dd") : null;
      const todayStr = Utilities.formatDate(today, "EST", "yyyy-MM-dd");
      let cachedData = [];
      if (cachedDateStr === todayStr && cachedHourlyData) {
        cachedData = JSON.parse(cachedHourlyData);
        Logger.log("getInventoryData: Using cached hourly data for today");
      }

      // Get the current time
      const currentHour = now.getHours();
      const currentTime = now.getTime() / 1000; // Convert to Unix timestamp

      // Use cached data for past hours (9:00 AM to current time)
      const startOfDay = new Date(today);
      startOfDay.setHours(9, 0, 0, 0); // Start at 9:00 AM today (changed from 10:00 AM)
      const startOfDayUnix = Math.floor(startOfDay.getTime() / 1000);

      for (let hour = 9; hour <= currentHour && hour <= 22; hour++) { // Start at 9:00 AM
        const hourTime = new Date(today);
        hourTime.setHours(hour, 0, 0, 0);
        const hourUnix = Math.floor(hourTime.getTime() / 1000);
        const cachedHour = cachedData.find(h => h.dt === hourUnix);
        if (cachedHour) {
          combinedHourlyData.push(cachedHour);
        } else {
          // Fallback if no cached data for this hour
          combinedHourlyData.push({
            dt: hourUnix,
            temp: data.overview.weather.temp,
            weather: data.overview.weather.description
          });
        }
      }

      // Use fresh data for future hours (current time to 10:00 PM)
      const endOfDay = new Date(today);
      endOfDay.setHours(22, 0, 0, 0); // 10:00 PM today
      const endOfDayUnix = Math.floor(endOfDay.getTime() / 1000);

      const futureData = freshHourlyData.filter(hour => {
        return hour.dt > currentTime && hour.dt <= endOfDayUnix;
      });

      combinedHourlyData = combinedHourlyData.concat(futureData);
      Logger.log(`getInventoryData: Combined hourly data - ${combinedHourlyData.length} entries`);
    } catch (e) {
      Logger.log(`getInventoryData: Error combining hourly weather data - ${e.message}`);
      combinedHourlyData = freshHourlyData;
    }

    data.overview.hourlyWeather = combinedHourlyData;

    var projectionSheet = spreadsheet.getSheetByName("D: Medley Projections");
    if (!projectionSheet) {
      Logger.log("getInventoryData: Sheet 'D: Medley Projections' not found");
      data.overview.salesProjections = [0, 0, 0, 0, 0, 0, 0];
      data.overview.salesOverrides = ['', '', '', '', '', '', ''];
    } else {
      data.overview.salesProjections = projectionSheet.getRange("B2:H2").getValues()[0];
      data.overview.salesOverrides = projectionSheet.getRange("B3:H3").getValues()[0];
      Logger.log("getInventoryData: Fetched sales projections");
    }

    var variableDataSheet = spreadsheet.getSheetByName("D: Variable Data");
    if (!variableDataSheet) {
      Logger.log("getInventoryData: Sheet 'D: Variable Data' not found");
      data.overview.productMix = { names: [], averages: [], overrides: [] };
    } else {
      data.overview.productMix = {
        names: variableDataSheet.getRange("A6:A13").getValues().flat(),
        averages: variableDataSheet.getRange("G6:G13").getValues().flat(),
        overrides: variableDataSheet.getRange("F6:F13").getValues().flat()
      };
      Logger.log("getInventoryData: Fetched product mix data");
    }

    Logger.log(`getInventoryData: Final data object summary - Overview date: ${data.overview.date}, Calendar events present: ${data.overview.calendarEvents ? 'Yes' : 'No'}`);
    return data;
  } catch (error) {
    Logger.log(`getInventoryData: Error - ${error.message}`);
    return {
      overview: {
        date: "Error",
        dayOfWeek: "Error",
        weather: { temp: "N/A", description: "Error fetching data" },
        calendarEvents: { todayEvents: [], upcomingEvents: [], ongoingEvents: [] },
        salesProjections: [0, 0, 0, 0, 0, 0, 0],
        salesOverrides: ['', '', '', '', '', '', ''],
        productMix: { names: [], averages: [], overrides: [] }
      },
      cookies: { names: [], amounts: [] },
      medleys: { names: [], inventory: [], catering: [] },
      sauces: { names: [], inventory: [], quarts: [] },
      garnishes: { names: [], inventory: [], quarts: [] },
      vegetables: { names: [], amounts: [] },
      dressings: { names: [], amounts: [] },
      cateringMedleys: { names: [], portions: [] }
    };
  }
}


function updateSalesOverrides(updates) {
  var projectionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("D: Medley Projections");
  updates.forEach(update => {
    projectionSheet.getRange("B3:H3").getCell(1, update.col).setValue(update.value);
  });
}

function updateData(updates) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("InventoryPrep");
  updates.forEach(update => {
    sheet.getRange(update.range).getCell(update.row, update.col).setValue(update.value);
  });
}

function getPrepSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("InventoryPrep");
  return sheet.getRange("A70:D139").getValues();
}

function processInventory(updates) {
  if (!Array.isArray(updates)) {
    throw new Error("Invalid updates provided to processInventory");
  }
  if (updates.length === 0) {
    return getPrepSheet();
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("InventoryPrep");
  updates.forEach(update => {
    sheet.getRange(update.range).getCell(update.row, update.col).setValue(update.value);
  });

  return getPrepSheet();
}

function ensureSettingsSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = spreadsheet.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = spreadsheet.insertSheet("Settings");
    settingsSheet.getRange("A1").setValue("Sales Projection Multiplier");
    settingsSheet.getRange("A2").setValue("Minimum Stock Threshold");
    settingsSheet.getRange("A3").setValue("Disable Validation"); // New setting
    settingsSheet.getRange("B1").setValue(1);
    settingsSheet.getRange("B2").setValue(0);
    settingsSheet.getRange("B3").setValue("No"); // Default to "No" (validation enabled)
    settingsSheet.getRange("D1").setValue("Authorized Users");
    settingsSheet.getRange("D2").setValue("admin@example.com");
    settingsSheet.getRange("D1:D1").setFontWeight("bold");
  }
  return settingsSheet;
}

function getSettings() {
  Logger.log("getSettings() function called");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var projectionSheet = spreadsheet.getSheetByName("D: Medley Projections");
  var detailedToDoSheet = spreadsheet.getSheetByName("Detailed To Do");
  var cookiesSheet = spreadsheet.getSheetByName("D: Cookies");
  var settingsSheet = spreadsheet.getSheetByName("Settings");
  
  var settings = {
    buffer: {
      cookies: 0,
      medleys: {Sunday: 0, Monday: 0, Tuesday: 0, Wednesday: 0, Thursday: 0, Friday: 0, Saturday: 0},
      catering: {medleyBuffer: 0},
      sauces: {bufferPar: [], bufferAmount: []},
      vegetables: {holdTimes: []}
    },
    authorizedUsers: []
  };
  
  if (cookiesSheet) {
    settings.buffer.cookies = cookiesSheet.getRange("C1").getValue() || 0;
  }
  
  if (projectionSheet) {
    settings.buffer.medleys.Sunday = projectionSheet.getRange("L39").getValue() || 0;
    settings.buffer.medleys.Monday = projectionSheet.getRange("L40").getValue() || 0;
    settings.buffer.medleys.Tuesday = projectionSheet.getRange("L41").getValue() || 0;
    settings.buffer.medleys.Wednesday = projectionSheet.getRange("L42").getValue() || 0;
    settings.buffer.medleys.Thursday = projectionSheet.getRange("L43").getValue() || 0;
    settings.buffer.medleys.Friday = projectionSheet.getRange("L44").getValue() || 0;
    settings.buffer.medleys.Saturday = projectionSheet.getRange("L45").getValue() || 0;
    settings.buffer.catering.medleyBuffer = projectionSheet.getRange("B50").getValue() || 0;
    
    var bufferParNames = projectionSheet.getRange("A5:A13").getValues().flat();
    var bufferParValues = projectionSheet.getRange("C5:C13").getValues().flat();
    settings.buffer.sauces.bufferPar = bufferParNames.map((name, index) => ({
      name: name || "Sauce " + (index + 1),
      daysWorth: bufferParValues[index] || 1
    }));
    
    var bufferAmountNames = projectionSheet.getRange("A16:A23").getValues().flat();
    var bufferAmountValues = projectionSheet.getRange("B16:B23").getValues().flat();
    settings.buffer.sauces.bufferAmount = bufferAmountNames.map((name, index) => ({
      name: name || "Sauce " + (index + 1),
      daysWorth: bufferAmountValues[index] || 1
    }));
  }
  
  if (detailedToDoSheet) {
    try {
      var vegData = detailedToDoSheet.getRange("A20:E40").getValues();
      for (let i = 1; i < vegData.length; i++) {
        if (vegData[i][0]) {
          settings.buffer.vegetables.holdTimes.push({
            name: vegData[i][0] || "Unknown",
            minDays: (typeof vegData[i][3] === 'number') ? vegData[i][3] : 1,
            maxDays: (typeof vegData[i][4] === 'number') ? vegData[i][4] : 7
          });
        }
      }
    } catch (e) {
      Logger.log("Error processing vegetables: " + e.message);
    }
  }
  
  if (settingsSheet) {
    settings.authorizedUsers = settingsSheet.getRange("D2:D10").getValues().flat().filter(Boolean);
  }
  
  return settings;
}

function addAuthorizedUser(email) {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  const users = settingsSheet.getRange("D2:D10").getValues().flat().filter(Boolean);
  
  if (users.includes(email)) {
    throw new Error("User already exists");
  }

  const newRow = users.length + 2;
  if (newRow > 10) {
    throw new Error("Maximum user limit reached (9 users)");
  }

  settingsSheet.getRange(`D${newRow}`).setValue(email);
  return settingsSheet.getRange("D2:D10").getValues().flat().filter(Boolean);
}

function removeAuthorizedUser(index) {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  const usersRange = settingsSheet.getRange("D2:D10");
  const users = usersRange.getValues().flat().filter(Boolean);

  if (index < 0 || index >= users.length) {
    throw new Error("Invalid user index");
  }

  users.splice(index, 1);
  usersRange.clear();
  if (users.length > 0) {
    settingsSheet.getRange("D2", 4, users.length, 1).setValues(users.map(u => [u]));
  }

  return users;
}

function saveSettings(updates) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var projectionSheet = spreadsheet.getSheetByName("D: Medley Projections");
  var detailedToDoSheet = spreadsheet.getSheetByName("Detailed To Do");
  var cookiesSheet = spreadsheet.getSheetByName("D: Cookies");

  if (!projectionSheet) throw new Error("Sheet 'D: Medley Projections' not found");
  if (!detailedToDoSheet) throw new Error("Sheet 'Detailed To Do' not found");
  if (!cookiesSheet) throw new Error("Sheet 'D: Cookies' not found");

  updates.forEach(update => {
    if (update.sheet === "D: Medley Projections") {
      projectionSheet.getRange(update.range).setValue(update.value);
    } else if (update.sheet === "Detailed To Do") {
      detailedToDoSheet.getRange(update.range).setValue(update.value);
    } else if (update.sheet === "D: Cookies") {
      cookiesSheet.getRange(update.range).setValue(update.value);
    }
  });
}

function getCateringOrders() {
  try {
    Logger.log("getCateringOrders: Starting catering orders fetch");
    const CATERING_SPREADSHEET_ID = "1HZNw30JH3oHpld1EZ8lybjP9GqpuoS9gewAHuo75J0w";
    
    let cateringSpreadsheet;
    try {
      cateringSpreadsheet = SpreadsheetApp.openById(CATERING_SPREADSHEET_ID);
      Logger.log("getCateringOrders: Accessed catering spreadsheet");
    } catch (e) {
      Logger.log("getCateringOrders: Error opening catering spreadsheet: " + e.message);
      return [];
    }
    
    const cateringSheet = cateringSpreadsheet.getSheetByName("Catering Orders");
    if (!cateringSheet) {
      Logger.log("getCateringOrders: Catering Orders sheet not found");
      return [];
    }
    
    const data = cateringSheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log("getCateringOrders: No catering orders found (only header row)");
      return [];
    }
    
    const headers = data[0];
    const findColumnIndex = (name) => headers.indexOf(name);
    
    const orderIdCol = findColumnIndex("Order ID");
    const dateCol = findColumnIndex("Date (YYYY-MM-DD)");
    const timeDueCol = findColumnIndex("Time Due");
    const deliveryAddressCol = findColumnIndex("Delivery Address");
    const deliveryInstructionsCol = findColumnIndex("Delivery Instructions");
    const customerNameCol = findColumnIndex("Customer Name");
    const customerEmailCol = findColumnIndex("Customer Email");
    const customerContactCol = findColumnIndex("Customer Contact");
    const itemsCol = findColumnIndex("Items (JSON)");
    const pdfLinkCol = findColumnIndex("PDF Link");
    const statusCol = findColumnIndex("Status");
    const distanceCol = findColumnIndex("Distance from Restaurant (Miles)");
    const travelTimeCol = findColumnIndex("Travel Time to Restaurant");
    
    if (dateCol === -1 || itemsCol === -1) {
      Logger.log("getCateringOrders: Required columns not found");
      return [];
    }
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const sevenDaysLater = new Date(today);
    sevenDaysLater.setDate(today.getDate() + 7);
    
    const orders = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[dateCol]) continue;
      
      let orderDate;
      try {
        orderDate = row[dateCol] instanceof Date ? new Date(row[dateCol]) : new Date(row[dateCol]);
        if (isNaN(orderDate.getTime())) continue;
        orderDate.setHours(0, 0, 0, 0);
      } catch (e) {
        Logger.log("getCateringOrders: Invalid date in row " + i + ": " + e.message);
        continue;
      }
      
      if (orderDate >= today && orderDate <= sevenDaysLater) {
        let itemsJson = { byog: [], bowls: [], drinks: [], additional: [] };
        try {
          if (row[itemsCol] && typeof row[itemsCol] === 'string' && row[itemsCol].trim() !== '') {
            itemsJson = JSON.parse(row[itemsCol]);
          }
        } catch (e) {
          Logger.log("getCateringOrders: Error parsing JSON in row " + i + ": " + e.message);
        }
        
        const order = {
          rowIndex: i,
          orderId: row[orderIdCol] || ("Order-" + i),
          date: Utilities.formatDate(orderDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          timeDue: row[timeDueCol] || "",
          deliveryAddress: row[deliveryAddressCol] || "",
          deliveryInstructions: row[deliveryInstructionsCol] || "",
          customerName: row[customerNameCol] || "Customer",
          customerEmail: row[customerEmailCol] || "",
          customerContact: row[customerContactCol] || "",
          items: itemsJson,
          pdfLink: row[pdfLinkCol] || "",
          status: row[statusCol] || "Pending",
          distance: row[distanceCol] || "",
          travelTime: row[travelTimeCol] || ""
        };
        orders.push(order);
      }
    }
    
    if (orders.length === 0) {
      Logger.log("getCateringOrders: No orders found for the next 7 days");
      return [];
    }
    
    const groupedOrders = {};
    orders.forEach(order => {
      const date = order.date;
      if (!groupedOrders[date]) groupedOrders[date] = [];
      groupedOrders[date].push(order);
    });
    
    const sortedDates = Object.keys(groupedOrders).sort();
    const result = sortedDates.map(date => {
      const dayOrders = groupedOrders[date];
      const formatted = formatDate(new Date(date));
      const dailyTotals = calculateDailyTotals(dayOrders);
      return { date, formattedDate: formatted, orders: dayOrders, totals: dailyTotals };
    });
    
    Logger.log(`getCateringOrders: Returning ${result.length} grouped orders`);
    return result;
  } catch (error) {
    Logger.log("getCateringOrders: Error - " + error.message);
    return [];
  }
}

function formatDate(date) {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${days[date.getDay()]}, ${months[date.getMonth()]} ${date.getDate()}`;
}

function calculateDailyTotals(orders) {
  const totals = {
    proteins: {}, bases: {}, gustos: {}, chips: {}, drinks: {}, desserts: {}, additionalItems: {}
  };
  
  if (!Array.isArray(orders)) return totals;
  
  const normalizeItemName = (name, category) => {
    if (!name) return category ? `Unknown ${category}` : "Unknown Item";
    name = name.trim();
    if (category === 'chips') {
      if (name.toLowerCase().includes("original")) return "Original Chips";
      if (name.toLowerCase().includes("bbq")) return "BBQ Chips";
      if (name.toLowerCase().includes("cinnamon")) return "Cinnamon Sugar Chips";
    }
    if (category === 'proteins') {
      if (name.toLowerCase().includes("tofu")) return "Umami Tofu";
      if (name.toLowerCase().includes("panko") || name.toLowerCase().includes("crispy")) return "Panko Crispy Chicken";
    }
    return name;
  };
  
  const normalizeServings = (servings) => {
    if (typeof servings !== 'number' || isNaN(servings)) return 1;
    return Math.round(servings * 10) / 10;
  };
  
  orders.forEach(order => {
    if (!order || !order.items) return;
    const items = order.items;
    
    if (items.byog && Array.isArray(items.byog)) {
      items.byog.forEach(byog => {
        if (!byog) return;
        if (byog.proteins) byog.proteins.forEach(protein => {
          if (protein && protein.name) {
            const name = normalizeItemName(protein.name, 'proteins');
            const servings = normalizeServings(protein.servings);
            totals.proteins[name] = (totals.proteins[name] || 0) + servings;
          }
        });
        if (byog.base_distribution) {
          Object.keys(byog.base_distribution).forEach(base => {
            if (base) {
              const servings = normalizeServings(byog.base_distribution[base]);
              totals.bases[base] = (totals.bases[base] || 0) + servings;
            }
          });
        }
        if (byog.gustos) byog.gustos.forEach(gusto => {
          if (gusto && gusto.name) {
            const name = normalizeItemName(gusto.name, 'gustos');
            const servings = normalizeServings(gusto.servings);
            totals.gustos[name] = (totals.gustos[name] || 0) + servings;
          }
        });
        if (byog.chips) byog.chips.forEach(chip => {
          if (chip && chip.name) {
            const name = normalizeItemName(chip.name, 'chips');
            const servings = normalizeServings(chip.servings);
            totals.chips[name] = (totals.chips[name] || 0) + servings;
          }
        });
        if (byog.adders) byog.adders.forEach(adder => {
          if (adder && adder.name) {
            const name = adder.name;
            const servings = normalizeServings(adder.servings);
            totals.additionalItems[name] = (totals.additionalItems[name] || 0) + servings;
          }
        });
      });
    }
    
    if (items.bowls && Array.isArray(items.bowls)) {
      items.bowls.forEach(bowl => {
        if (!bowl) return;
        const quantity = bowl.quantity || 1;
        if (bowl.gustos) bowl.gustos.forEach(gusto => {
          if (gusto && gusto.name) {
            const name = normalizeItemName(gusto.name, 'gustos');
            const servings = normalizeServings(gusto.servings || quantity);
            totals.gustos[name] = (totals.gustos[name] || 0) + servings;
          }
        });
        if (bowl.bases) bowl.bases.forEach(base => {
          if (base && base.name) {
            const servings = normalizeServings(base.servings || quantity);
            totals.bases[base.name] = (totals.bases[base.name] || 0) + servings;
          }
        });
        if (bowl.proteins) bowl.proteins.forEach(protein => {
          if (protein && protein.name) {
            const name = normalizeItemName(protein.name, 'proteins');
            const servings = normalizeServings(protein.servings || quantity);
            totals.proteins[name] = (totals.proteins[name] || 0) + servings;
          }
        });
        if (bowl.chips) bowl.chips.forEach(chip => {
          if (chip && chip.name) {
            const name = normalizeItemName(chip.name, 'chips');
            const servings = normalizeServings(chip.servings || quantity);
            totals.chips[name] = (totals.chips[name] || 0) + servings;
          }
        });
        if (bowl.adders) bowl.adders.forEach(adder => {
          if (adder && adder.name) {
            const servings = normalizeServings(adder.servings || quantity);
            totals.additionalItems[adder.name] = (totals.additionalItems[adder.name] || 0) + servings;
          }
        });
      });
    }
    
    if (items.drinks && Array.isArray(items.drinks)) {
      items.drinks.forEach(drink => {
        if (!drink) return;
        const quantity = drink.quantity || 1;
        let name = drink.itemName || "Unspecified Drink";
        if (drink.flavor) name += ` (${drink.flavor})`;
        else if (drink.sweetness) name += ` (${drink.sweetness})`;
        totals.drinks[name] = (totals.drinks[name] || 0) + quantity;
      });
    }
    
    if (items.additional && Array.isArray(items.additional)) {
      items.additional.forEach(item => {
        if (!item) return;
        const quantity = item.quantity || 1;
        const name = item.itemName || "Additional Item";
        if (name.toLowerCase().includes("cookie") || 
            name.toLowerCase().includes("brownie") || 
            name.toLowerCase().includes("blondie") || 
            name.toLowerCase().includes("marshmallow")) {
          totals.desserts[name] = (totals.desserts[name] || 0) + quantity;
        } else {
          totals.additionalItems[name] = (totals.additionalItems[name] || 0) + quantity;
        }
      });
    }
  });
  
  return {
    proteins: objectToSortedArray(totals.proteins),
    bases: objectToSortedArray(totals.bases),
    gustos: objectToSortedArray(totals.gustos),
    chips: objectToSortedArray(totals.chips),
    drinks: objectToSortedArray(totals.drinks),
    desserts: objectToSortedArray(totals.desserts),
    additionalItems: objectToSortedArray(totals.additionalItems)
  };
}

function objectToSortedArray(obj) {
  if (!obj || typeof obj !== 'object') return [];
  return Object.entries(obj)
    .map(([name, servings]) => ({ name, servings }))
    .sort((a, b) => b.servings - a.servings);
}

function updateCateringOrderStatus(orderId, rowIndex, newStatus) {
  try {
    const CATERING_SPREADSHEET_ID = "1HZNw30JH3oHpld1EZ8lybjP9GqpuoS9gewAHuo75J0w";
    const cateringSpreadsheet = SpreadsheetApp.openById(CATERING_SPREADSHEET_ID);
    const cateringSheet = cateringSpreadsheet.getSheetByName("Catering Orders");
    
    if (!cateringSheet) throw new Error("Catering Orders sheet not found");
    
    const headers = cateringSheet.getRange(1, 1, 1, cateringSheet.getLastColumn()).getValues()[0];
    const statusColIndex = headers.indexOf("Status") + 1;
    if (statusColIndex === 0) throw new Error("Status column not found");
    
    cateringSheet.getRange(rowIndex + 1, statusColIndex).setValue(newStatus);
    return { success: true };
  } catch (error) {
    Logger.log(`Error in updateCateringOrderStatus: ${error.message}`);
    return { error: error.message };
  }
}

function deleteCateringOrder(orderId, rowIndex) {
  try {
    const CATERING_SPREADSHEET_ID = "1HZNw30JH3oHpld1EZ8lybjP9GqpuoS9gewAHuo75J0w";
    const cateringSpreadsheet = SpreadsheetApp.openById(CATERING_SPREADSHEET_ID);
    const cateringSheet = cateringSpreadsheet.getSheetByName("Catering Orders");
    
    if (!cateringSheet) throw new Error("Catering Orders sheet not found");
    cateringSheet.deleteRow(rowIndex + 1);
    return { success: true };
  } catch (error) {
    Logger.log(`Error in deleteCateringOrder: ${error.message}`);
    return { error: error.message };
  }
}

function getCalendarEvents() {
  Logger.log("getCalendarEvents: Starting event fetch");
  const calendarIds = [
    { id: 'c_c59c18554960ff4ecb8488d624cac17f8af62224f3c71acd42edefb4bd52ea6a@group.calendar.google.com', source: 'general' },
    { id: 'marketing@whatsyourgusto.com', source: 'marketing' }
  ];
  
  const today = new Date();
  const timeZone = Session.getScriptTimeZone();
  today.setHours(0, 0, 0, 0);
  const todayInTimeZone = new Date(Utilities.formatDate(today, timeZone, 'yyyy-MM-dd'));
  
  const nextWeek = new Date(todayInTimeZone.getTime() + 7 * 24 * 60 * 60 * 1000);
  const yesterday = new Date(todayInTimeZone.getTime() - 24 * 60 * 60 * 1000);
  
  Logger.log(`getCalendarEvents: Today (normalized): ${todayInTimeZone}, Next week: ${nextWeek}, Yesterday: ${yesterday}`);
  
  let allEvents = {
    todayEvents: [],
    upcomingEvents: [],
    ongoingEvents: []
  };
  
  calendarIds.forEach(calendar => {
    try {
      const calendarObj = CalendarApp.getCalendarById(calendar.id);
      if (!calendarObj) {
        Logger.log(`getCalendarEvents: Calendar not found or access denied: ${calendar.id}`);
        return;
      }
      
      const events = calendarObj.getEvents(yesterday, nextWeek);
      const calendarTimeZone = calendarObj.getTimeZone();
      
      Logger.log(`getCalendarEvents: Fetched ${events.length} events from calendar: ${calendar.id}, Source: ${calendar.source}`);
      
      events.forEach(event => {
        const title = event.getTitle();
        let start, end;
        
        const startDate = new Date(event.getStartTime());
        const endDate = new Date(event.getEndTime());
        
        Logger.log(`getCalendarEvents: Event: ${title}, Start: ${startDate}, End: ${endDate}, Source: ${calendar.source}`);
        
        if (event.isAllDayEvent()) {
          start = Utilities.formatDate(startDate, calendarTimeZone, 'MMMM dd');
          const adjustedEndDate = new Date(endDate.getTime() - 24 * 60 * 60 * 1000);
          end = Utilities.formatDate(adjustedEndDate, calendarTimeZone, 'MMMM dd');
        } else {
          start = Utilities.formatDate(startDate, calendarTimeZone, 'MMMM dd, h:mm a');
          end = Utilities.formatDate(endDate, calendarTimeZone, 'MMMM dd, h:mm a');
        }
        
        const startDateOnly = new Date(Utilities.formatDate(startDate, calendarTimeZone, 'yyyy-MM-dd'));
        const endDateOnly = new Date(Utilities.formatDate(endDate, calendarTimeZone, 'yyyy-MM-dd'));
        
        Logger.log(`getCalendarEvents: Event ${title} - StartDateOnly: ${startDateOnly}, EndDateOnly: ${endDateOnly}, Today: ${todayInTimeZone}`);
        
        const eventData = {
          title,
          start,
          end,
          startDate: startDate.toISOString(),
          endDate: endDate.toISOString(),
          source: calendar.source // Ensure source is set
        };
        
        // Categorize events with priority
        if (startDateOnly.getTime() === todayInTimeZone.getTime()) {
          allEvents.todayEvents.push(eventData);
          Logger.log(`getCalendarEvents: Event ${title} categorized as Going Live Today!`);
        } else if (startDateOnly.getTime() < todayInTimeZone.getTime() && endDateOnly.getTime() >= todayInTimeZone.getTime()) {
          allEvents.ongoingEvents.push(eventData);
          Logger.log(`getCalendarEvents: Event ${title} categorized as Ongoing Events/Promos`);
        } else if (startDateOnly.getTime() > todayInTimeZone.getTime() && startDateOnly.getTime() <= nextWeek.getTime()) {
          allEvents.upcomingEvents.push(eventData);
          Logger.log(`getCalendarEvents: Event ${title} categorized as Upcoming Events/Promos`);
        }
      });
    } catch (error) {
      Logger.log(`getCalendarEvents: Error fetching events from calendar ${calendar.id} - ${error.message}`);
    }
  });
  
  allEvents.todayEvents.sort((a, b) => new Date(a.startDate) - new Date(b.startDate));
  allEvents.upcomingEvents.sort((a, b) => new Date(a.startDate) - new Date(b.startDate));
  allEvents.ongoingEvents.sort((a, b) => new Date(a.startDate) - new Date(b.startDate));
  
  Logger.log(`getCalendarEvents: Final event counts - Today: ${allEvents.todayEvents.length}, Upcoming: ${allEvents.upcomingEvents.length}, Ongoing: ${allEvents.ongoingEvents.length}`);
  
  return allEvents;
}

// Expose the function to the client-side
function getCalendarEventsForWebApp() {
  return getCalendarEvents();
}

function testFunction() {
  Logger.log("testFunction: Executing test function");
  return {
    test: "This is a test response",
    timestamp: new Date().toISOString()
  };
}
