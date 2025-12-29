/**
 * ------------------------------------------------------------------
 * 1. MENU & UI SETUP
 * ------------------------------------------------------------------
 */

const PROPS_KEYS = {
  CALENDAR_ID: "calendarId",
  COLOR_MAPPINGS: "colorMappings",
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⏰ Timetracking")
    .addItem("📊 Open Dashboard", "showSidebar")
    .addSeparator()
    .addItem("🔄 Sync Last Month (Manual Trigger)", "importLastMonthAuto")
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("page")
    .setTitle("📆 Timetracking Settings")
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * ------------------------------------------------------------------
 * 2. AUTOMATION TRIGGERS
 * ------------------------------------------------------------------
 */

// Run this on the 1st of every month via Apps Script Triggers
function importLastMonthAuto() {
  const calendarId = getSavedCalendarId();
  if (!calendarId) {
    console.error("No Calendar ID saved. Run the Dashboard first.");
    return;
  }

  const now = new Date();
  // Start: 1st day of previous month
  const startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  // End: 1st day of current month (exclusive)
  const endDate = new Date(now.getFullYear(), now.getMonth(), 1);

  coreImportEvents(calendarId, startDate, endDate);
}

/**
 * ------------------------------------------------------------------
 * 3. CORE IMPORT LOGIC
 * ------------------------------------------------------------------
 */

function importEventsFromSidebar(data) {
  if (!data.calendarId) return { success: false, message: "Missing Calendar ID" };
  
  saveCalendarId(data.calendarId);

  // Parse dates from sidebar
  const startDate = new Date(data.startDate);
  const endDate = new Date(data.endDate);
  endDate.setDate(endDate.getDate() + 1); // Make inclusive

  return coreImportEvents(data.calendarId, startDate, endDate, data.searchText);
}

function coreImportEvents(calendarId, startDate, endDate, searchText) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const header = [["Title", "Description", "Start", "End", "Duration (hrs)", "Project"]];
    
    // Add header only if sheet is blank
    if (sheet.getLastRow() === 0) {
      sheet.getRange("A1:F1").setValues(header).setFontWeight("bold");
    }

    const events = [];
    let pageToken;
    const userEmail = Session.getActiveUser().getEmail();

    // -- FETCH EVENTS --
    do {
      const response = Calendar.Events.list(calendarId, {
        timeMin: startDate.toISOString(),
        timeMax: endDate.toISOString(),
        singleEvents: true,
        orderBy: "startTime",
        q: searchText || undefined,
        pageToken: pageToken,
      });

      if (!response.items) break;

      const filtered = response.items.filter(event => {
        // 1. Filter Declined
        if (event.attendees) {
          const self = event.attendees.find(a => a.email === userEmail || a.self);
          if (self && self.responseStatus === "declined") return false;
        }
        // 2. Filter All-Day Events (User Request)
        if (event.start.date) return false; 
        
        return true;
      });

      events.push(...filtered);
      pageToken = response.nextPageToken;
    } while (pageToken);

    if (events.length === 0) return { success: false, message: "No events found." };

    // -- PROCESS DATA --
    const savedMappings = PropertiesService.getUserProperties().getProperty(PROPS_KEYS.COLOR_MAPPINGS);
    const projectMappings = savedMappings ? JSON.parse(savedMappings) : {};

    const eventRows = events.map(event => {
      const start = new Date(event.start.dateTime);
      const end = new Date(event.end.dateTime);
      const duration = (end - start) / (1000 * 60 * 60);

      // Map Colors
      const colorId = event.colorId || "Default";
      let projectName = projectMappings[colorId];
      
      // Fallback for unmapped
      if (!projectName) {
        projectName = (colorId === "Default") ? "Internal / Non-Client" : "Unmapped Color";
      }

      return [
        event.summary || "(No Title)",
        cleanHtmlText(event.description || ""),
        start,
        end,
        duration,
        projectName,
      ];
    });

    // -- CLEAR & WRITE --
    // Safe clear: don't touch headers
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, header[0].length).clearContent();
    }

    const outputRange = sheet.getRange(2, 1, eventRows.length, header[0].length);
    outputRange.setValues(eventRows);

    // Formats
    sheet.getRange(2, 3, eventRows.length, 2).setNumberFormat("dd/MM/yyyy HH:mm");
    sheet.getRange(2, 5, eventRows.length, 1).setNumberFormat("0.00");
    sheet.autoResizeColumns(1, header[0].length);

    return { success: true };

  } catch (error) {
    if (error.toString().includes("Calendar is not defined")) {
      return { success: false, message: "❌ Enable 'Google Calendar API' in Services." };
    }
    return { success: false, message: error.toString() };
  }
}

/**
 * ------------------------------------------------------------------
 * 4. SETTINGS & HELPERS
 * ------------------------------------------------------------------
 */

function getSettings() {
  // We need the ID to fetch samples. 
  // If not saved, guess the user's email.
  const calendarId = getSavedCalendarId() || Session.getActiveUser().getEmail();
  
  return {
    calendarId: calendarId,
    // Pass the ID so we can look up real events
    colors: getCalendarColors(calendarId) 
  };
}

function updateProjectMapping(colorId, projectName) {
  const props = PropertiesService.getUserProperties();
  const mappings = JSON.parse(props.getProperty(PROPS_KEYS.COLOR_MAPPINGS) || "{}");
  
  mappings[colorId] = projectName;
  props.setProperty(PROPS_KEYS.COLOR_MAPPINGS, JSON.stringify(mappings));
  
  return { success: true };
}

function getCalendarColors(calendarId) {
  // 1. Define Standard Palette
  const standardColors = {
    "Default": "#ffffff", 
    "1": "#7986cb", "2": "#33b679", "3": "#8e24aa", "4": "#e67c73",
    "5": "#f6c026", "6": "#f4511e", "7": "#039be5", "8": "#616161",
    "9": "#3f51b5", "10": "#0b8043", "11": "#d50000"
  };

  // 2. Fetch Saved Mappings
  const props = PropertiesService.getUserProperties();
  const mappings = JSON.parse(props.getProperty(PROPS_KEYS.COLOR_MAPPINGS) || "{}");

  // 3. Find Sample Events (The missing feature)
  const samples = {};
  
  try {
    // Look back 3 months to find representative events
    const now = new Date();
    const threeMonthsAgo = new Date();
    threeMonthsAgo.setMonth(now.getMonth() - 3);

    // Use Advanced API for speed
    const response = Calendar.Events.list(calendarId, {
      timeMin: threeMonthsAgo.toISOString(),
      timeMax: now.toISOString(),
      singleEvents: true,
      maxResults: 250, // Don't need infinite events, just enough to get samples
      orderBy: "startTime"
    });

    if (response.items) {
      // Loop through events and grab the first title found for each color
      response.items.forEach(event => {
        const cId = event.colorId || "Default";
        // Only save the sample if we haven't found one for this color yet
        if (!samples[cId] && event.summary) {
          samples[cId] = event.summary;
        }
      });
    }
  } catch (e) {
    console.warn("Could not fetch samples: " + e.toString());
    // We continue even if this fails, just without samples
  }

  // 4. Merge Data for Frontend
  const colorList = [];
  for (const [id, hex] of Object.entries(standardColors)) {
    colorList.push({
      id: id,
      hex: hex,
      name: mappings[id] || "",
      // Attach the sample event title we found, or a fallback
      sample: samples[id] || "(No recent events with this color)"
    });
  }
  
  return colorList;
}

function cleanHtmlText(html) {
  if (!html) return "";
  let text = html.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<[^>]+>/g, "");
  text = text.replace(/&nbsp;/g, " ");
  return text.trim();
}

function getSavedCalendarId() {
  return PropertiesService.getUserProperties().getProperty(PROPS_KEYS.CALENDAR_ID);
}

function saveCalendarId(id) {
  PropertiesService.getUserProperties().setProperty(PROPS_KEYS.CALENDAR_ID, id);
}