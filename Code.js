function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⏰ Timetracking")
    .addItem("📊 Open Dashboard", "showSidebar")
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("page")
    .setTitle("Settings")
    .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Properties Service keys
const PROPS_KEYS = {
  CALENDAR_ID: "calendarId",
  COLOR_MAPPINGS: "colorMappings",
};

// Helper function to clean HTML from text
function cleanHtmlText(html) {
  if (!html) return "";

  let text = html;

  // Replace <br>, <br/>, and <br /> tags with newlines first
  text = text.replace(/<br\s*\/?>/gi, "\n");

  // Replace common HTML entities
  const htmlEntities = {
    "&nbsp;": " ",
    "&amp;": "&",
    "&lt;": "<",
    "&gt;": ">",
    "&quot;": '"',
    "&apos;": "'",
    "&#39;": "'",
    "&ndash;": "–",
    "&mdash;": "—",
    "&bull;": "•",
  };

  // Replace HTML entities with their text equivalents
  Object.entries(htmlEntities).forEach(([entity, replacement]) => {
    text = text.replace(new RegExp(entity, "g"), replacement);
  });

  // Remove all remaining HTML tags
  text = text.replace(/<[^>]+>/g, "");

  // Decode any numeric HTML entities
  text = text.replace(/&#(\d+);/g, function (match, dec) {
    return String.fromCharCode(dec);
  });

  // Clean up whitespace
  text = text.replace(/\s+/g, " ").trim();

  return text;
}

// Get calendar color information with saved project mappings
function getCalendarColors(calendarId) {
  try {
    if (!calendarId) return {};

    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) return {};

    // Get all events for the last 3 months to analyze colors
    const threeMonthsAgo = new Date();
    threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);
    const events = calendar.getEvents(threeMonthsAgo, new Date());

    // Get saved project mappings
    const savedMappings = PropertiesService.getUserProperties().getProperty(
      PROPS_KEYS.COLOR_MAPPINGS
    );
    const projectMappings = savedMappings ? JSON.parse(savedMappings) : {};

    // Create a map of colors and their associated event titles
    const colorMap = {};
    events.forEach((event) => {
      const color = event.getColor();
      if (color && !colorMap[color]) {
        colorMap[color] = {
          colorId: color,
          colorHex: getColorHexById(color),
          sampleEvent: event.getTitle(),
          projectName: projectMappings[color] || "",
        };
      }
    });

    return colorMap;
  } catch (error) {
    console.error("Error getting calendar colors:", error);
    return {};
  }
}

function getColorHexById(id) {
  const colors = {
    1: "#a4bdfc",
    2: "#7ae7bf",
    3: "#dbadff",
    4: "#ff887c",
    5: "#fbd75b",
    6: "#ffb878",
    7: "#46d6db",
    8: "#e1e1e1",
    9: "#5484ed",
    10: "#51b749",
    11: "#dc2127",
  };
  return colors[id] || "#ffffff";
}

// Update project name for a color
function updateProjectMapping(colorId, projectName) {
  const userProperties = PropertiesService.getUserProperties();
  const savedMappings = userProperties.getProperty(PROPS_KEYS.COLOR_MAPPINGS);
  const projectMappings = savedMappings ? JSON.parse(savedMappings) : {};

  projectMappings[colorId] = projectName;
  userProperties.setProperty(
    PROPS_KEYS.COLOR_MAPPINGS,
    JSON.stringify(projectMappings)
  );

  return getCalendarColors(getSavedCalendarId());
}

// Save and retrieve calendar ID
function getSavedCalendarId() {
  return PropertiesService.getUserProperties().getProperty(
    PROPS_KEYS.CALENDAR_ID
  );
}

function saveCalendarId(calendarId) {
  PropertiesService.getUserProperties().setProperty(
    PROPS_KEYS.CALENDAR_ID,
    calendarId
  );
}

// Import events from sidebar
function importEventsFromSidebar(data) {
  try {
    // Validate inputs
    if (!data.calendarId) {
      return { success: false, message: "Please provide a valid Calendar ID." };
    }

    const startDate = new Date(data.startDate);
    const endDate = new Date(data.endDate);

    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      return {
        success: false,
        message: "Please provide valid start and end dates.",
      };
    }

    // Save calendar ID for future use
    saveCalendarId(data.calendarId);

    // Get calendar
    const calendar = CalendarApp.getCalendarById(data.calendarId);
    if (!calendar) {
      return {
        success: false,
        message: "The Calendar ID provided is invalid or inaccessible.",
      };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Set header
    const header = [
      ["Title", "Description", "Start", "End", "Duration (hrs)", "Project"],
    ];
    sheet.getRange("A1:F1").setValues(header).setFontWeight("bold");

    // Fetch events
    const events = data.searchText
      ? calendar.getEvents(startDate, endDate, { search: data.searchText })
      : calendar.getEvents(startDate, endDate);

    if (events.length === 0) {
      return {
        success: false,
        message: "No events found for the specified criteria.",
      };
    }

    // Get saved project mappings
    const savedMappings = PropertiesService.getUserProperties().getProperty(
      PROPS_KEYS.COLOR_MAPPINGS
    );
    const projectMappings = savedMappings ? JSON.parse(savedMappings) : {};

    // Prepare data for output
    const eventRows = events.map((event) => {
      const duration =
        (event.getEndTime() - event.getStartTime()) / (1000 * 60 * 60); // Convert ms to hours
      return [
        event.getTitle(),
        cleanHtmlText(event.getDescription()),
        event.getStartTime(),
        event.getEndTime(),
        duration,
        projectMappings[event.getColor()] || "–",
      ];
    });

    // Clear existing data (except header)
    const lastRow = Math.max(sheet.getLastRow(), 2);
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, header[0].length).clear();
    }

    // Write events to sheet
    const outputRange = sheet.getRange(
      2,
      1,
      eventRows.length,
      header[0].length
    );
    outputRange.setValues(eventRows);

    // Format date and duration columns
    sheet
      .getRange(2, 3, eventRows.length, 2)
      .setNumberFormat("dd/MM/yyyy HH:mm");
    sheet.getRange(2, 5, eventRows.length, 1).setNumberFormat("0.00");

    // Auto-resize columns
    sheet.autoResizeColumns(1, header[0].length);

    // Add summary at the bottom
    const totalRow = eventRows.length + 2;
    sheet.getRange(totalRow, 1, 1, 4).merge().setValue("Total Hours:");
    sheet.getRange(totalRow, 5).setFormula(`=SUM(E2:E${eventRows.length + 1})`);
    sheet.getRange(totalRow, 5).setNumberFormat("0.00");

    return { success: true };
  } catch (error) {
    console.error(error);
    return { success: false, message: error.toString() };
  }
}
