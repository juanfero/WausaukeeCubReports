  // ================================
  // HOTEL DAILY REPORTS SYSTEM - VERSIÓN COMPLETA Y ACTUALIZADA
  // ================================

  // MAIN MENU CREATION
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Wausaukee Club Reports')
    .addItem('Generate Meal Report', 'generateMealReport')
    .addItem('Tomorrow\'s Meal Report', 'tomorrowMealReport')
    .addSeparator()
    .addItem('Generate Occupancy Report', 'generateOccupancyReport')
    .addItem('Tomorrow\'s Occupancy Report', 'tomorrowOccupancyReport')
    .addSeparator()
    // NUEVO: Menú de Housekeeping
    .addSubMenu(ui.createMenu(' Housekeeping Reports')
        .addItem('Generate Housekeeping Report', 'generateHousekeepingReport')
        .addItem('Today\'s Housekeeping Report', 'todaysHousekeepingReport')
        .addSeparator()
        .addItem('Generate Housekeeping Outlook', 'generateHousekeepingOutlook'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Outlook Reports')
      .addItem('Weekly Outlook (7 days)', 'generateWeeklyOutlook')
      .addItem('Bi-Weekly Outlook (14 days)', 'generateBiWeeklyOutlook')
      .addItem('Monthly Outlook (30 days)', 'generateMonthlyOutlook')
      .addItem('Custom Period Outlook', 'generateCustomOutlook')
      .addSeparator() // Separador para las nuevas opciones
      .addItem('Generate Meal Outlook', 'generateMealOutlook') // NUEVA OPCIÓN
      .addItem('Generate Occupancy Outlook', 'generateOccupancyOutlook')) // NUEVA OPCIÓN
    .addSeparator()
    .addSubMenu(ui.createMenu('Airport Transfers')
      .addItem(' All Upcoming Transfers', 'generateUpcomingTransfers')
      .addItem('Refresh Transfer Schedule', 'generateUpcomingTransfers'))
    .addSeparator()
    .addItem(' Cleanup Old Reports', 'cleanupOldReports')
    .addToUi();
}

  // ================================
  // DAILY MEAL REPORT FUNCTIONS
  // ================================

  // MAIN FUNCTION: Generate meal report with date picker
  function generateMealReport() {
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.prompt(
      'Generate Meal Report',
      'Enter the date for the report inYYYY-MM-DD format\n(Example: 2025-05-22):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const dateInput = response.getResponseText().trim();
      
      if (!dateInput) {
        ui.alert('Error', 'Please enter a valid date.', ui.ButtonSet.OK);
        return;
      }
      
      try {
        ui.alert(
          'Processing...',
          `Generating meal report for ${dateInput}...`,
          ui.ButtonSet.OK
        );
        
        processMealReport(dateInput);
        
        ui.alert(
          'Report Generated Successfully',
          `The meal report for ${dateInput} has been generated successfully.\n\nYou can find it in the "Meal Report" sheet.\n\n(Previous report data has been replaced with the new date)`,
          ui.ButtonSet.OK
        );
        
      } catch (error) {
        ui.alert(
          'Error',
          `Could not generate the report:\n\n${error.message}\n\nPlease verify:\n• Date format isYYYY-MM-DD\n• "Original Data" sheet exists\n• There is data for that date`,
          ui.ButtonSet.OK
        );
        debugLog('Error in meal report:', error.message);
      }
    }
  }

  // QUICK FUNCTION: Today's meal report
  function tomorrowMealReport() {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);

  const formattedDate = `${tomorrow.getFullYear()}-${(tomorrow.getMonth() + 1).toString().padStart(2, '0')}-${tomorrow.getDate().toString().padStart(2, '0')}`;
  
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert(
      'Processing...',
      `Generating meal report for tomorrow (${formattedDate})...`,
      ui.ButtonSet.OK
    );
    
    processMealReport(formattedDate);
    
    ui.alert(
      'Report Generated Successfully',
      `Tomorrow's meal report (${formattedDate}) has been generated successfully.\n\nYou can find it in the "Meal Report" sheet.\n\n(Previous report data has been replaced with tomorrow's data)`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert(
      'Error',
      `Could not generate the meal report:\n\n${error.message}\n\nPlease verify that the "Original Data" sheet exists with valid information.`,
      ui.ButtonSet.OK
    );
    debugLog('Error in tomorrow\'s report:', error.message);
  }
}


  // FUNCTION: Process daily meal report
  function processMealReport(dateText) {
    // Parse date
    const dateParts = dateText.split('-');
    if (dateParts.length !== 3) {
      throw new Error('Invalid date format. UseYYYY-MM-DD');
    }
    
    const year = parseInt(dateParts[0]);
    const month = parseInt(dateParts[1]) - 1;
    const day = parseInt(dateParts[2]);
    
    if (isNaN(year) || isNaN(month) || isNaN(day)) {
      throw new Error('Invalid date. Use valid numbers.');
    }
    
    const date = new Date(year, month, day);
    
    // Get sheets
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const originalData = ss.getSheetByName('Original Data');
    
    if (!originalData) {
      throw new Error('Could not find "Original Data" sheet');
    }
    
    // Get daily data
    const dailyData = getDailyData(originalData, date);
    
    // Usar una sola hoja fija "Meal Report"
    const fixedSheetName = 'Meal Report';
    let reportSheet = ss.getSheetByName(fixedSheetName);
    
    if (reportSheet) {
      reportSheet.clear();
    } else {
      reportSheet = ss.insertSheet(fixedSheetName);
    }
    
    // Generate report
    generateDailyReport(reportSheet, dailyData, date);
  }

  // FUNCTION: Generate daily report in sheet (UPDATED FOR MEAL SUMMARY DETAIL)
  // FUNCTION: Generate daily report in sheet (UPDATED FOR NEW MEAL SUMMARY)
// FUNCTION: Generate daily report in sheet (UPDATED FOR NEW MEAL SUMMARY & ALPHABETICAL SORTING)
// FUNCTION: Generate daily report in sheet (UPDATED to include Customer ID)
function generateDailyReport(sheet, dailyData, date) {
    sheet.clear();

    const formattedDate = formatReadableDate(date);
    sheet.getRange(1, 1).setValue(`DAILY MEAL REPORT - ${formattedDate}`);
    // El reporte ahora tiene 10 columnas de detalle
    sheet.getRange(1, 1, 1, 10).merge().setHorizontalAlignment('center');
    sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');

    // ===== NUEVO ENCABEZADO =====
    const detailHeaders = [
        'Cabin', 'Cabin Name', 'Customer ID', 'Guest Name', 'Type', 'Check-in', 'Check-out', 'Breakfast', 'Lunch', 'Dinner'
    ];
    sheet.getRange(3, 1, 1, 10).setValues([detailHeaders]);

    // Detail data
    const detailRows = [];
    dailyData.guests.forEach(guest => {
        const dailyMeals = calculateDailyMeals(guest, date);
        // ===== NUEVA FILA CON CUSTOMER ID =====
        detailRows.push([
            guest.cabinNumber,
            guest.cabinName,
            guest.customerID, 
            guest.guestName,
            guest.guestType,
            formatDate(guest.checkIn),
            formatDate(guest.checkOut),
            dailyMeals.breakfast,
            dailyMeals.lunch,
            dailyMeals.dinner
        ]);
    });

    if (detailRows.length > 0) {
        sheet.getRange(4, 1, detailRows.length, 10).setValues(detailRows);
        // El rango de comidas ahora empieza en la columna 8
        const mealRange = sheet.getRange(4, 8, detailRows.length, 3);
        mealRange.setNumberFormat('0');
    }

    // --- La sección de Summary no se modifica y sigue funcionando igual ---
    const summaryStartRow = detailRows.length + 6;
    sheet.getRange(summaryStartRow, 1).setValue(`MEAL SUMMARY - ${formatReadableDate(date)}`);
    sheet.getRange(summaryStartRow, 1, 1, 5).merge().setHorizontalAlignment('center');
    sheet.getRange(summaryStartRow, 1).setFontWeight('bold').setFontSize(14);

    const summaryHeaders = ['Cabin', 'Cabin Name', 'Breakfast', 'Lunch', 'Dinner'];
    sheet.getRange(summaryStartRow + 2, 1, 1, 5).setValues([summaryHeaders]);

    const cabinSummary = {};
    dailyData.guests.forEach(guest => {
        const cabinKey = guest.cabinNumber;
        if (!cabinSummary[cabinKey]) {
            cabinSummary[cabinKey] = {
                cabinName: guest.cabinName,
                breakfastTotal: 0, breakfastChildren: 0,
                lunchTotal: 0, lunchChildren: 0,
                dinnerTotal: 0, dinnerChildren: 0
            };
        }
        const dailyMeals = calculateDailyMeals(guest, date);
        if (dailyMeals.breakfast === 1) {
            cabinSummary[cabinKey].breakfastTotal++;
            if (guest.guestType === 'Child') cabinSummary[cabinKey].breakfastChildren++;
        }
        if (dailyMeals.lunch === 1) {
            cabinSummary[cabinKey].lunchTotal++;
            if (guest.guestType === 'Child') cabinSummary[cabinKey].lunchChildren++;
        }
        if (dailyMeals.dinner === 1) {
            cabinSummary[cabinKey].dinnerTotal++;
            if (guest.guestType === 'Child') cabinSummary[cabinKey].dinnerChildren++;
        }
    });

    const cabinSummaryArray = Object.keys(cabinSummary).map(key => ({
        cabinNumber: key, ...cabinSummary[key]
    }));
    cabinSummaryArray.sort((a, b) => a.cabinName.localeCompare(b.cabinName));

    const summaryRows = cabinSummaryArray.map(cabin => {
        const breakfastText = `Total: ${cabin.breakfastTotal} (Children: ${cabin.breakfastChildren})`;
        const lunchText = `Total: ${cabin.lunchTotal} (Children: ${cabin.lunchChildren})`;
        const dinnerText = `Total: ${cabin.dinnerTotal} (Children: ${cabin.dinnerChildren})`;
        return [cabin.cabinNumber, cabin.cabinName, breakfastText, lunchText, dinnerText];
    });

    if (summaryRows.length > 0) {
        sheet.getRange(summaryStartRow + 3, 1, summaryRows.length, 5).setValues(summaryRows);
    }

    const totalRow = summaryStartRow + 3 + summaryRows.length + 1;
    let grandTotalBreakfast = 0, grandTotalLunch = 0, grandTotalDinner = 0;
    dailyData.guests.forEach(guest => {
        const dailyMeals = calculateDailyMeals(guest, date);
        if (dailyMeals.breakfast) grandTotalBreakfast++;
        if (dailyMeals.lunch) grandTotalLunch++;
        if (dailyMeals.dinner) grandTotalDinner++;
    });

    sheet.getRange(totalRow, 1, 1, 5).setValues([
        ['TOTAL GENERAL', '', `Total: ${grandTotalBreakfast}`, `Total: ${grandTotalLunch}`, `Total: ${grandTotalDinner}`]
    ]);

    formatDailySheet(sheet, detailRows.length, summaryStartRow);
    formatDailySummary(sheet, summaryStartRow, summaryRows.length, totalRow);
}
  // FORMAT FUNCTIONS (DAILY MEAL REPORT)
  // FORMAT FUNCTIONS (DAILY MEAL REPORT)
function formatDailySheet(sheet, numDetailRows, summaryStartRow) {
    // Ajustado a 10 columnas
    const headerRange = sheet.getRange(3, 1, 1, 10);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#34a853');
    headerRange.setFontColor('white');
    
    if (numDetailRows > 0) {
        // Ajustado a 10 columnas
        const detailRange = sheet.getRange(3, 1, numDetailRows + 1, 10);
        detailRange.setBorder(true, true, true, true, true, true);
        
        // El rango de comidas ahora empieza en la columna 8
        const mealRange = sheet.getRange(4, 8, numDetailRows, 3);
        mealRange.setHorizontalAlignment('center');
        mealRange.setNumberFormat('0');
        
        // La columna 'Type' ahora es la 5
        const typeRange = sheet.getRange(4, 5, numDetailRows, 1);
        typeRange.setHorizontalAlignment('center');
        
        alternateCabinColors(sheet, numDetailRows, 4);
    }
    
    sheet.autoResizeColumns(1, 10);
}

  // UPDATED FORMATTING for the new Meal Summary
function formatDailySummary(sheet, startRow, numRows, totalRow) {
    // Adjust ranges to 5 columns instead of 9
    const headerRange = sheet.getRange(startRow + 2, 1, 1, 5);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#ff9800');
    headerRange.setFontColor('white');

    if (numRows > 0) {
        const summaryRange = sheet.getRange(startRow + 3, 1, numRows, 5);
        summaryRange.setBorder(true, true, true, true, true, true);
        summaryRange.setBackground('#fff3e0');

        // Center the meal columns
        const mealColumns = sheet.getRange(startRow + 3, 3, numRows, 3);
        mealColumns.setHorizontalAlignment('center');
    }

    const totalRange = sheet.getRange(totalRow, 1, 1, 5);
    totalRange.setFontWeight('bold');
    totalRange.setBackground('#f57c00');
    totalRange.setFontColor('white');
    totalRange.setBorder(true, true, true, true, true, true);
    
    const totalMealRange = sheet.getRange(totalRow, 3, 1, 3);
    totalMealRange.setHorizontalAlignment('center');
}

  // ================================
  // DAILY OCCUPANCY REPORT FUNCTIONS
  // ================================
  // MAIN FUNCTION: Generate occupancy report with date picker
  function generateOccupancyReport() {
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.prompt(
      'Generate Occupancy Report',
      'Enter the date for the report in YYYY-MM-DD format\n(Example: 2025-05-22):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const dateInput = response.getResponseText().trim();
      
      if (!dateInput) {
        ui.alert('Error', 'Please enter a valid date.', ui.ButtonSet.OK);
        return;
      }
      
      try {
        ui.alert(
          'Processing...',
          `Generating occupancy report for ${dateInput}...`,
          ui.ButtonSet.OK
        );
        
        processOccupancyReport(dateInput);
        
        ui.alert(
          'Report Generated Successfully',
          `The occupancy report for ${dateInput} has been generated successfully.\n\nYou can find it in the "Occupancy" sheet.`,
          ui.ButtonSet.OK
        );
        
      } catch (error) {
        ui.alert(
          'Error',
          `Could not generate the report:\n\n${error.message}\n\nPlease verify:\n• Date format is YYYY-MM-DD\n• "Original Data" sheet exists\n• There is data for that date`,
          ui.ButtonSet.OK
        );
        debugLog('Error in occupancy report:', error.message);
      }
    }
  }

  // QUICK FUNCTION: Today's occupancy report
  function tomorrowOccupancyReport() {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);

  const formattedDate = `${tomorrow.getFullYear()}-${(tomorrow.getMonth() + 1).toString().padStart(2, '0')}-${tomorrow.getDate().toString().padStart(2, '0')}`;
  
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert(
      'Processing...',
      `Generating occupancy report for tomorrow (${formattedDate})...`,
      ui.ButtonSet.OK
    );
    
    processOccupancyReport(formattedDate);
    
    ui.alert(
      'Report Generated Successfully',
      `Tomorrow's occupancy report (${formattedDate}) has been generated successfully.\n\nYou can find it in the "Occupancy" sheet.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert(
      'Error',
      `Could not generate the occupancy report:\n\n${error.message}\n\nPlease verify that the "Original Data" sheet exists with valid information.`,
      ui.ButtonSet.OK
    );
    debugLog('Error in tomorrow\'s occupancy report:', error.message);
  }
}


  // FUNCTION: Process occupancy report
  function processOccupancyReport(dateText) {
    // Parse date
    const dateParts = dateText.split('-');
    if (dateParts.length !== 3) {
      throw new Error('Invalid date format. Use YYYY-MM-DD');
    }
    
    const year = parseInt(dateParts[0]);
    const month = parseInt(dateParts[1]) - 1;
    const day = parseInt(dateParts[2]);
    
    if (isNaN(year) || isNaN(month) || isNaN(day)) {
      throw new Error('Invalid date. Use valid numbers.');
    }
    
    const date = new Date(year, month, day);
    
    // Get sheets
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const originalData = ss.getSheetByName('Original Data');
    
    if (!originalData) {
      throw new Error('Could not find "Original Data" sheet');
    }
    
    // Get occupancy data
    const occupancyData = getOccupancyData(originalData, date); // <--- ESTA LLAMADA AHORA FUNCIONARÁ
    
    // Create or get occupancy sheet
    let occupancySheet = ss.getSheetByName('Occupancy');
    if (!occupancySheet) {
      occupancySheet = ss.insertSheet('Occupancy');
    }
    
    // Generate report
    generateOccupancyReportInSheet(occupancySheet, occupancyData, date);
  }

  
  // FUNCTION: Get occupancy data for a specific date
  function getOccupancyData(dataSheet, queryDate) {
    const data = dataSheet.getDataRange().getValues();
    const headers = data[0];
    
    const cleanDate = new Date(queryDate.getFullYear(), queryDate.getMonth(), queryDate.getDate());
    
    // Load cabin names
    const cabinNames = getCabinNames();
    
    // Find columns
    const colJOB = headers.indexOf('JOB ID');
    const colFullName = headers.indexOf('FullName');
    const colArrivalDate = headers.indexOf('ArrivalDate');
    const colDepartDate = headers.indexOf('DepartDate');
    const colItemID = headers.indexOf('Item ID');
    
    if (colJOB === -1 || colFullName === -1 || colArrivalDate === -1 || colDepartDate === -1) {
      throw new Error('Could not find all required columns');
    }
    
    if (colItemID === -1) {
      throw new Error('Could not find "Item ID" column. Please verify the column exists.');
    }
    
    const occupiedCabins = new Map();
    const duplicatesDetected = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const jobId = row[colJOB];
      const fullName = row[colFullName];
      const arrivalDateRaw = row[colArrivalDate];
      const departDateRaw = row[colDepartDate];
      const itemID = row[colItemID];
      
      if (!jobId || !fullName || !arrivalDateRaw || !departDateRaw) continue;
      
      if (!itemID || !itemID.toString().trim().toUpperCase().startsWith('FB')) {
        continue;
      }
      
      const uniqueKey = `${jobId}|${fullName}|${arrivalDateRaw}|${departDateRaw}`;
      if (duplicatesDetected.has(uniqueKey)) {
        continue;
      }
      duplicatesDetected.add(uniqueKey);
      
      let arrivalDate, departureDate;
      
      if (typeof arrivalDateRaw === 'string') {
        arrivalDate = parseAmericanDate(arrivalDateRaw);
      } else {
        arrivalDate = new Date(arrivalDateRaw);
      }
      
      if (typeof departDateRaw === 'string') {
        departureDate = parseAmericanDate(departDateRaw);
      } else {
        departureDate = new Date(departDateRaw);
      }
      
      if (!arrivalDate || !departureDate || isNaN(arrivalDate.getTime()) || isNaN(departureDate.getTime())) {
        continue;
      }
      
      const cleanArrival = new Date(arrivalDate.getFullYear(), arrivalDate.getMonth(), arrivalDate.getDate());
      const cleanDeparture = new Date(departureDate.getFullYear(), departureDate.getMonth(), departureDate.getDate());
      
      const isOccupied = cleanArrival <= cleanDate && cleanDeparture >= cleanDate;
      
      if (isOccupied) {
        const jobIdStr = jobId.toString().trim();
        const cabinName = cabinNames[jobIdStr] || `Cabin ${jobIdStr}`;
        const guestType = getGuestType(itemID);
        
        if (!occupiedCabins.has(jobIdStr)) {
          occupiedCabins.set(jobIdStr, {
            cabinNumber: jobId,
            cabinName: cabinName,
            arrivalDate: arrivalDate,
            departureDate: departureDate,
            guests: [],
            adults: 0,
            children: 0
          });
        }
        
        const cabin = occupiedCabins.get(jobIdStr);
        cabin.guests.push(`${fullName}`); // Modificado para no incluir (Adult/Child) y que coincida con el reporte
        
        if (guestType === 'Adult') {
          cabin.adults++;
        } else if (guestType === 'Child') {
          cabin.children++;
        }
      }
    }
    
    const cabinsArray = Array.from(occupiedCabins.values());
    cabinsArray.sort((a, b) => {
      if (a.cabinNumber && b.cabinNumber) {
          return a.cabinNumber.localeCompare(b.cabinNumber);
      }
      return 0;
    });
    
    return cabinsArray;
  }


  // FUNCTION: Generate occupancy report in sheet
  function generateOccupancyReportInSheet(sheet, occupancyData, date) {
    sheet.clear();
    
    // Main title
    const formattedDate = formatReadableDate(date);
    sheet.getRange(1, 1).setValue(`CABIN OCCUPANCY REPORT - ${formattedDate}`);
    sheet.getRange(1, 1, 1, 8).merge().setHorizontalAlignment('center');
    sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setBackground('#2196f3').setFontColor('white');
    
    // Headers
    const headers = [
      'Cabin Number', 'Cabin Name', 'Arrival Date', 'Departure Date', 'Adults', 'Children', 'Total', 'Guests'
    ];
    sheet.getRange(3, 1, 1, 8).setValues([headers]);
    
    // Occupancy data
    const detailRows = [];
    let totalGuests = 0;
    let totalAdults = 0;
    let totalChildren = 0;
    
    occupancyData.forEach(cabin => {
      const guestsList = cabin.guests.join(', ');
      const totalCabinGuests = cabin.adults + cabin.children;
      
      totalGuests += totalCabinGuests;
      totalAdults += cabin.adults;
      totalChildren += cabin.children;
      
      detailRows.push([
        cabin.cabinNumber,
        cabin.cabinName,
        formatDate(cabin.arrivalDate),
        formatDate(cabin.departureDate),
        cabin.adults,
        cabin.children,
        totalCabinGuests,
        guestsList
      ]);
    });
    
    // Write data
    if (detailRows.length > 0) {
      sheet.getRange(4, 1, detailRows.length, 8).setValues(detailRows);
    }
    
    // Summary
    const summaryRow = detailRows.length + 6;
    sheet.getRange(summaryRow, 1).setValue(`OCCUPANCY SUMMARY`);
    sheet.getRange(summaryRow, 1, 1, 8).merge().setHorizontalAlignment('center');
    sheet.getRange(summaryRow, 1).setFontWeight('bold').setFontSize(14);
    
    sheet.getRange(summaryRow + 2, 1, 1, 8).setValues([
      ['Total Occupied Cabins', occupancyData.length, '', '', totalAdults, totalChildren, totalGuests, '']
    ]);
    
    // Format
    formatOccupancySheet(sheet, detailRows.length, summaryRow);
  }

  // FORMAT FUNCTIONS (DAILY OCCUPANCY REPORT)
  function formatOccupancySheet(sheet, numRows, summaryRow) {
    // Format headers
    const headerRange = sheet.getRange(3, 1, 1, 8);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#2196f3');
    headerRange.setFontColor('white');
    
    if (numRows > 0) {
      // Format data
      const detailRange = sheet.getRange(3, 1, numRows + 1, 8);
      detailRange.setBorder(true, true, true, true, true, true);
      
      // Center specific columns
      sheet.getRange(4, 1, numRows, 1).setHorizontalAlignment('center');
      sheet.getRange(4, 3, numRows, 2).setHorizontalAlignment('center');
      sheet.getRange(4, 5, numRows, 3).setHorizontalAlignment('center');
      
      // Alternate colors
      for (let i = 0; i < numRows; i++) {
        if (i % 2 === 1) {
          sheet.getRange(4 + i, 1, 1, 8).setBackground('#f8f9fa');
        }
      }
    }
    
    // Format summary
    const summaryRange = sheet.getRange(summaryRow + 2, 1, 1, 8);
    summaryRange.setFontWeight('bold');
    summaryRange.setBackground('#4caf50');
    summaryRange.setFontColor('white');
    summaryRange.setHorizontalAlignment('center');
    summaryRange.setBorder(true, true, true, true, true, true);
    
    // Adjust columns
    sheet.autoResizeColumns(1, 8);
    sheet.setColumnWidth(8, 400);
  }

  // ================================
  // OUTLOOK REPORTS FUNCTIONS (Combined for Boss)
  // ================================

  // FUNCTION: Generate Weekly Outlook (7 days)
  function generateWeeklyOutlook() {
    generateOutlookReport(7, 'Weekly');
  }

  // FUNCTION: Generate Bi-Weekly Outlook (14 days)
  function generateBiWeeklyOutlook() {
    generateOutlookReport(14, 'Bi-Weekly');
  }

  // FUNCTION: Generate Monthly Outlook (30 days)
  function generateMonthlyOutlook() {
    generateOutlookReport(30, 'Monthly');
  }

  // FUNCTION: Generate Custom Period Outlook
  function generateCustomOutlook() {
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.prompt(
      'Custom Period Outlook',
      'Enter the number of days for the outlook report\n(Example: 21 for 3 weeks):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const daysInput = response.getResponseText().trim();
      const days = parseInt(daysInput);
      
      if (isNaN(days) || days < 1 || days > 365) {
        ui.alert('Error', 'Please enter a valid number of days (1-365).', ui.ButtonSet.OK);
        return;
      }
      
      generateOutlookReport(days, `${days}-Day`);
    }
  }

  // MAIN FUNCTION: Generate Combined Outlook Report (for boss)
  function generateOutlookReport(numberOfDays, periodType) {
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.prompt(
      `${periodType} Outlook Report`,
      'Enter the START date for the outlook period inYYYY-MM-DD format\n(Example: 2025-06-08):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const dateInput = response.getResponseText().trim();
      
      if (!dateInput) {
        ui.alert('Error', 'Please enter a valid start date.', ui.ButtonSet.OK);
        return;
      }
      
      try {
        ui.alert(
          'Processing...',
          `Generating ${periodType.toLowerCase()} outlook report for ${numberOfDays} days starting ${dateInput}...`,
          ui.ButtonSet.OK
        );
        
        processOutlookReport(dateInput, numberOfDays, periodType);
        
        ui.alert(
          'Report Generated Successfully',
          `The ${periodType.toLowerCase()} outlook report has been generated successfully.\n\nYou can find it in the "Outlook" sheet.\n\nThis report shows meal and occupancy projections for ${numberOfDays} days starting from ${dateInput}.`,
          ui.ButtonSet.OK
        );
        
      } catch (error) {
        ui.alert(
          'Error',
          `Could not generate the outlook report:\n\n${error.message}\n\nPlease verify:\n• Date format isYYYY-MM-DD\n• "Original Data" sheet exists\n• There is data for the specified period`,
          ui.ButtonSet.OK
        );
        debugLog('Error in outlook report:', error.message);
      }
    }
  }

  
  // FUNCTION: Process combined outlook report
function processOutlookReport(startDateText, numberOfDays, periodType) {
  const dateParts = startDateText.split('-');
  if (dateParts.length !== 3) {
    throw new Error('Invalid date format. Use YYYY-MM-DD');
  }
  const year = parseInt(dateParts[0]);
  const month = parseInt(dateParts[1]) - 1;
  const day = parseInt(dateParts[2]);
  if (isNaN(year) || isNaN(month) || isNaN(day)) {
    throw new Error('Invalid date. Use valid numbers.');
  }
  const startDate = new Date(year, month, day);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const originalData = ss.getSheetByName('Original Data');
  if (!originalData) {
    throw new Error('Could not find "Original Data" sheet');
  }
  
  const dateRange = [];
  for (let i = 0; i < numberOfDays; i++) {
    const currentDate = new Date(startDate);
    currentDate.setDate(startDate.getDate() + i);
    dateRange.push(currentDate);
  }
  
  const outlookData = getOutlookData(originalData, dateRange);
  
  let outlookSheet = ss.getSheetByName('Outlook');
  if (!outlookSheet) {
    outlookSheet = ss.insertSheet('Outlook');
  }
  
  // CORRECCIÓN: Se verifica que la llamada a la función incluya todos los parámetros necesarios.
  generateOutlookReportInSheet(outlookSheet, outlookData, dateRange, periodType);
}


// FUNCTION: Generate combined outlook report in sheet
function generateOutlookReportInSheet(sheet, outlookData, dateRange, periodType) {
  sheet.clear();
  
  // Se agrega una validación para evitar errores si el rango de fechas es inválido
  if (!dateRange || dateRange.length === 0) {
    sheet.getRange(1,1).setValue('Error: Invalid date range provided.');
    return;
  }
  
  const startDate = formatReadableDate(dateRange[0]);
  const endDate = formatReadableDate(dateRange[dateRange.length - 1]);
  
  sheet.getRange(1, 1).setValue(`${periodType.toUpperCase()} OUTLOOK REPORT`);
  sheet.getRange(1, 1, 1, 9).merge().setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold').setBackground('#673ab7').setFontColor('white');
  
  sheet.getRange(2, 1).setValue(`Period: ${startDate} to ${endDate}`);
  sheet.getRange(2, 1, 1, 9).merge().setHorizontalAlignment('center').setFontSize(12).setFontWeight('bold');
  
  // MEAL PROJECTIONS SECTION
  sheet.getRange(4, 1).setValue('MEAL PROJECTIONS');
  sheet.getRange(4, 1, 1, 9).merge().setHorizontalAlignment('center').setFontWeight('bold').setFontSize(14).setBackground('#ff9800').setFontColor('white');
  
  const mealHeaders = ['Date', 'Day', 'Breakfast', 'Lunch', 'Dinner', 'Adults', 'Children', 'Total Guests', 'Summary'];
  sheet.getRange(6, 1, 1, 9).setValues([mealHeaders]);
  
  const mealRows = [];
  let totalBreakfast = 0, totalLunch = 0, totalDinner = 0, totalAdultsCombined = 0, totalChildrenCombined = 0, totalGuestsCombined = 0;
  
  outlookData.forEach(dayData => {
    const dayOfWeek = dayData.date.toLocaleDateString('en-US', { weekday: 'short' });
    const dateStr = formatDate(dayData.date);
    const summary = `B=${dayData.breakfast}, L=${dayData.lunch}, D=${dayData.dinner}`;
    totalBreakfast += dayData.breakfast;
    totalLunch += dayData.lunch;
    totalDinner += dayData.dinner;
    totalAdultsCombined += dayData.adults;
    totalChildrenCombined += dayData.children;
    totalGuestsCombined += dayData.totalOccupancy;
    mealRows.push([dateStr, dayOfWeek, dayData.breakfast, dayData.lunch, dayData.dinner, dayData.adults, dayData.children, dayData.totalOccupancy, summary]);
  });
  
  if (mealRows.length > 0) sheet.getRange(7, 1, mealRows.length, 9).setValues(mealRows);
  
  const mealTotalRow = 7 + mealRows.length + 1;
  sheet.getRange(mealTotalRow, 1, 1, 9).setValues([['TOTAL', '', totalBreakfast, totalLunch, totalDinner, totalAdultsCombined, totalChildrenCombined, totalGuestsCombined, `Total Period: ${totalBreakfast + totalLunch + totalDinner} meals`]]);
  
  // OCCUPANCY PROJECTIONS BY CABIN SECTION
  const occupancyStartRow = mealTotalRow + 3;
  sheet.getRange(occupancyStartRow, 1).setValue('OCCUPANCY PROJECTIONS BY CABIN');
  sheet.getRange(occupancyStartRow, 1, 1, 8).merge().setHorizontalAlignment('center').setFontWeight('bold').setFontSize(14).setBackground('#2196f3').setFontColor('white');
  
  const occupancyHeaders = ['Date', 'Day', 'Cabin', 'Cabin Name', 'Adults', 'Children', 'Arrival', 'Departure'];
  sheet.getRange(occupancyStartRow + 2, 1, 1, 8).setValues([occupancyHeaders]);
  
  const occupancyRows = [];
  outlookData.forEach(dayData => {
    const dayOfWeek = dayData.date.toLocaleDateString('en-US', { weekday: 'short' });
    const dateStr = formatDate(dayData.date);
    
    if (dayData.cabinBreakdown && dayData.cabinBreakdown.length > 0) {
      dayData.cabinBreakdown.forEach((cabin, index) => {
        // CORRECCIÓN: Se asegura que cada push sea un array (una fila)
        occupancyRows.push([
          index === 0 ? dateStr : '',
          index === 0 ? dayOfWeek : '',
          cabin.cabinNumber,
          cabin.cabinName,
          cabin.adults,
          cabin.children,
          formatDate(cabin.arrivalDate),
          formatDate(cabin.departureDate)
        ]);
      });
    } else {
      // CORRECCIÓN: Se asegura que este push también sea un array (una fila)
      occupancyRows.push([dateStr, dayOfWeek, 'No Occupancy', '', 0, 0, '', '']);
    }
  });
  
  if (occupancyRows.length > 0) sheet.getRange(occupancyStartRow + 3, 1, occupancyRows.length, 8).setValues(occupancyRows);
  
  const occupancyTotalRow = occupancyStartRow + 3 + occupancyRows.length + 1;
  const avgOccupancy = outlookData.length > 0 ? Math.round(totalGuestsCombined / outlookData.length) : 0;
  const avgAdults = outlookData.length > 0 ? Math.round(totalAdultsCombined / outlookData.length) : 0;
  const avgChildren = outlookData.length > 0 ? Math.round(totalChildrenCombined / outlookData.length) : 0;
  
  const totalsData = [[
    'TOTALS', '', `${totalAdultsCombined + totalChildrenCombined} Total Guests`, '', totalAdultsCombined, totalChildrenCombined, '', ''
  ]];
  const averagesData = [[
    'AVERAGES', '', `${avgOccupancy} Avg/Day`, '', avgAdults, avgChildren, '', ''
  ]];
  
  sheet.getRange(occupancyTotalRow, 1, 1, 8).setValues(totalsData);
  sheet.getRange(occupancyTotalRow + 1, 1, 1, 8).setValues(averagesData);
  
  formatOutlookSheet(sheet, mealRows.length, occupancyRows.length, mealTotalRow, occupancyStartRow, occupancyTotalRow);
}

  // FORMAT FUNCTIONS (COMBINED OUTLOOK REPORT)
  function formatOutlookSheet(sheet, mealRowCount, occupancyRowCount, mealTotalRow, occupancyStartRow, occupancyTotalRow) {
    const mealHeaderRange = sheet.getRange(6, 1, 1, 9);
    mealHeaderRange.setFontWeight('bold');
    mealHeaderRange.setBackground('#ff9800');
    mealHeaderRange.setFontColor('white');
    
    if (mealRowCount > 0) {
      const mealDataRange = sheet.getRange(6, 1, mealRowCount + 1, 9);
      mealDataRange.setBorder(true, true, true, true, true, true);
      
      sheet.getRange(7, 3, mealRowCount, 6).setHorizontalAlignment('center');
      sheet.getRange(7, 3, mealRowCount, 6).setNumberFormat('0');
      
      for (let i = 0; i < mealRowCount; i++) {
        if (i % 2 === 1) {
          sheet.getRange(7 + i, 1, 1, 9).setBackground('#fff3e0');
        }
      }
    }
    
    const mealTotalRange = sheet.getRange(mealTotalRow, 1, 1, 9);
    mealTotalRange.setFontWeight('bold');
    mealTotalRange.setBackground('#f57c00');
    mealTotalRange.setFontColor('white');
    mealTotalRange.setBorder(true, true, true, true, true, true);
    mealTotalRange.setHorizontalAlignment('center');
    
    const occupancyHeaderRange = sheet.getRange(occupancyStartRow + 2, 1, 1, 6);
    occupancyHeaderRange.setFontWeight('bold');
    occupancyHeaderRange.setBackground('#2196f3');
    occupancyHeaderRange.setFontColor('white');
    
    if (occupancyRowCount > 0) {
      const occupancyDataRange = sheet.getRange(occupancyStartRow + 2, 1, occupancyRowCount + 1, 6);
      occupancyDataRange.setBorder(true, true, true, true, true, true);
      
      sheet.getRange(occupancyStartRow + 3, 3, occupancyRowCount, 1).setHorizontalAlignment('center');
      sheet.getRange(occupancyStartRow + 3, 5, occupancyRowCount, 2).setHorizontalAlignment('center');
      sheet.getRange(occupancyStartRow + 3, 5, occupancyRowCount, 2).setNumberFormat('0');
      
      for (let i = 0; i < occupancyRowCount; i++) {
        if (i % 2 === 1) {
          sheet.getRange(occupancyStartRow + 3 + i, 1, 1, 6).setBackground('#e3f2fd');
        }
        if (i > 0 && sheet.getRange(occupancyStartRow + 3 + i, 1).getValue() === '') {
          const prevRowDate = sheet.getRange(occupancyStartRow + 3 + i - 1, 1).getValue();
          if (prevRowDate !== '') {
            let mergeCount = 0;
            for (let j = i; j >= 0; j--) {
              if (sheet.getRange(occupancyStartRow + 3 + j, 1).getValue() === '') {
                mergeCount++;
              } else {
                mergeCount++;
                break;
              }
            }
            if (mergeCount > 1) {
              sheet.getRange(occupancyStartRow + 3 + i - mergeCount + 1, 1, mergeCount, 1).mergeVertically();
              sheet.getRange(occupancyStartRow + 3 + i - mergeCount + 1, 2, mergeCount, 1).mergeVertically();
            }
          }
        }
      }
    }
    
    const occupancyTotalRange = sheet.getRange(occupancyTotalRow, 1, 2, 6);
    occupancyTotalRange.setFontWeight('bold');
    occupancyTotalRange.setBackground('#1976d2');
    occupancyTotalRange.setFontColor('white');
    occupancyTotalRange.setBorder(true, true, true, true, true, true);
    occupancyTotalRange.setHorizontalAlignment('center');
    
    sheet.autoResizeColumns(1, 9);
    sheet.setColumnWidth(4, 150);
    sheet.setColumnWidth(9, 200);
  }

  // ================================
  // NEW SEPARATED OUTLOOK REPORT FUNCTIONS
  // ================================

  // NUEVA FUNCIÓN: Generar Meal Outlook con selector de fecha
  function generateMealOutlook() {
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.prompt(
      'Generate Meal Outlook Report',
      'Enter the START date for the Meal Outlook period inYYYY-MM-DD format\n(Example: 2025-06-08):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const dateInput = response.getResponseText().trim();
      
      if (!dateInput) {
        ui.alert('Error', 'Please enter a valid start date.', ui.ButtonSet.OK);
        return;
      }
      
      try {
        const daysResponse = ui.prompt(
          'Generate Meal Outlook Report',
          'Enter the number of days for the Meal Outlook report (e.g., 7 for a week):',
          ui.ButtonSet.OK_CANCEL
        );

        if (daysResponse.getSelectedButton() !== ui.Button.OK) {
          return; // User cancelled
        }
        const numberOfDays = parseInt(daysResponse.getResponseText().trim());
        if (isNaN(numberOfDays) || numberOfDays < 1 || numberOfDays > 365) {
          ui.alert('Error', 'Please enter a valid number of days (1-365).', ui.ButtonSet.OK);
          return;
        }

        ui.alert(
          'Processing...',
          `Generating Meal Outlook report for ${numberOfDays} days starting ${dateInput}...`,
          ui.ButtonSet.OK
        );
        
        processMealOutlookReport(dateInput, numberOfDays);
        
        ui.alert(
          'Report Generated Successfully',
          `The Meal Outlook report has been generated successfully.\n\nYou can find it in the "Meal Outlook" sheet.`,
          ui.ButtonSet.OK
        );
        
      } catch (error) {
        ui.alert(
          'Error',
          `Could not generate the Meal Outlook report:\n\n${error.message}\n\nPlease verify:\n• Date format isYYYY-MM-DD\n• "Original Data" sheet exists`,
          ui.ButtonSet.OK
        );
        debugLog('Error in Meal Outlook report:', error.message);
      }
    }
  }

  // NUEVA FUNCIÓN: Generar Occupancy Outlook con selector de fecha
  function generateOccupancyOutlook() {
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.prompt(
      'Generate Occupancy Outlook Report',
      'Enter the START date for the Occupancy Outlook period inYYYY-MM-DD format\n(Example: 2025-06-08):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const dateInput = response.getResponseText().trim();
      
      if (!dateInput) {
        ui.alert('Error', 'Please enter a valid start date.', ui.ButtonSet.OK);
        return;
      }
      
      try {
        const daysResponse = ui.prompt(
          'Generate Occupancy Outlook Report',
          'Enter the number of days for the Occupancy Outlook report (e.g., 7 for a week):',
          ui.ButtonSet.OK_CANCEL
        );

        if (daysResponse.getSelectedButton() !== ui.Button.OK) {
          return; // User cancelled
        }
        const numberOfDays = parseInt(daysResponse.getResponseText().trim());
        if (isNaN(numberOfDays) || numberOfDays < 1 || numberOfDays > 365) {
          ui.alert('Error', 'Please enter a valid number of days (1-365).', ui.ButtonSet.OK);
          return;
        }

        ui.alert(
          'Processing...',
          `Generating Occupancy Outlook report for ${numberOfDays} days starting ${dateInput}...`,
          ui.ButtonSet.OK
        );
        
        processOccupancyOutlookReport(dateInput, numberOfDays);
        
        ui.alert(
          'Report Generated Successfully',
          `The Occupancy Outlook report has been generated successfully.\n\nYou can find it in the "Occupancy Outlook" sheet.`,
          ui.ButtonSet.OK
        );
        
      } catch (error) {
        ui.alert(
          'Error',
          `Could not generate the Occupancy Outlook report:\n\n${error.message}\n\nPlease verify:\n• Date format isYYYY-MM-DD\n• "Original Data" sheet exists`,
          ui.ButtonSet.OK
        );
        debugLog('Error in Occupancy Outlook report:', error.message);
      }
    }
  }

  // NUEVA FUNCIÓN: Procesa el reporte de Meal Outlook
  function processMealOutlookReport(startDateText, numberOfDays) {
    const dateParts = startDateText.split('-');
    if (dateParts.length !== 3) {
      throw new Error('Invalid date format. UseYYYY-MM-DD');
    }
    const year = parseInt(dateParts[0]);
    const month = parseInt(dateParts[1]) - 1;
    const day = parseInt(dateParts[2]);
    if (isNaN(year) || isNaN(month) || isNaN(day)) {
      throw new Error('Invalid date. Use valid numbers.');
    }
    const startDate = new Date(year, month, day);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const originalData = ss.getSheetByName('Original Data');
    if (!originalData) {
      throw new Error('Could not find "Original Data" sheet');
    }
    
    const dateRange = [];
    for (let i = 0; i < numberOfDays; i++) {
      const currentDate = new Date(startDate);
      currentDate.setDate(startDate.getDate() + i);
      dateRange.push(currentDate);
    }
    
    const outlookData = getOutlookData(originalData, dateRange);
    
    let mealOutlookSheet = ss.getSheetByName('Meal Outlook');
    if (!mealOutlookSheet) {
      mealOutlookSheet = ss.insertSheet('Meal Outlook');
    }
    
    generateMealOutlookInSheet(mealOutlookSheet, outlookData, dateRange, "Meal");
  }

  // NUEVA FUNCIÓN: Procesa el reporte de Occupancy Outlook
  // NUEVA FUNCIÓN: Procesa el reporte de Occupancy Outlook
function processOccupancyOutlookReport(startDateText, numberOfDays) {
  const dateParts = startDateText.split('-');
  if (dateParts.length !== 3) {
    throw new Error('Invalid date format. Use YYYY-MM-DD');
  }
  const year = parseInt(dateParts[0]);
  const month = parseInt(dateParts[1]) - 1;
  const day = parseInt(dateParts[2]);
  if (isNaN(year) || isNaN(month) || isNaN(day)) {
    throw new Error('Invalid date. Use valid numbers.');
  }
  const startDate = new Date(year, month, day);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const originalData = ss.getSheetByName('Original Data');
  if (!originalData) {
    throw new Error('Could not find "Original Data" sheet');
  }
  
  const dateRange = [];
  for (let i = 0; i < numberOfDays; i++) {
    const currentDate = new Date(startDate);
    currentDate.setDate(startDate.getDate() + i);
    dateRange.push(currentDate);
  }
  
  // 1. Obtener datos de Ocupación
  const outlookData = getOutlookData(originalData, dateRange);
  
  // 2. Obtener datos de Housekeeping
  const housekeepingSheet = ss.getSheetByName('Housekeeping');
  // Si la hoja no existe, se crea un mapa vacío para no generar errores.
  const housekeepingData = housekeepingSheet ? getHousekeepingDataForRange(housekeepingSheet, dateRange) : new Map();
  
  let occupancyOutlookSheet = ss.getSheetByName('Occupancy Outlook');
  if (!occupancyOutlookSheet) {
    occupancyOutlookSheet = ss.insertSheet('Occupancy Outlook');
  }
  
  // 3. Se pasan AMBOS grupos de datos a la función que dibuja el reporte
  generateOccupancyOutlookInSheet(occupancyOutlookSheet, outlookData, housekeepingData, dateRange, "Occupancy");
}

  // NUEVA FUNCIÓN: Genera el reporte de Meal Outlook en la hoja
  function generateMealOutlookInSheet(sheet, outlookData, dateRange, periodType) {
    sheet.clear();
    
    const startDate = formatReadableDate(dateRange[0]);
    const endDate = formatReadableDate(dateRange[dateRange.length - 1]);
    
    sheet.getRange(1, 1).setValue(`${periodType.toUpperCase()} OUTLOOK REPORT`);
    sheet.getRange(1, 1, 1, 9).merge().setHorizontalAlignment('center');
    sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setBackground('#673ab7').setFontColor('white');
    
    sheet.getRange(2, 1).setValue(`Period: ${startDate} to ${endDate}`);
    sheet.getRange(2, 1, 1, 9).merge().setHorizontalAlignment('center');
    sheet.getRange(2, 1).setFontSize(12).setFontWeight('bold');
    
    // MEAL PROJECTIONS SECTION (Solo esta sección)
    sheet.getRange(4, 1).setValue('MEAL PROJECTIONS');
    sheet.getRange(4, 1, 1, 9).merge().setHorizontalAlignment('center');
    sheet.getRange(4, 1).setFontWeight('bold').setFontSize(14).setBackground('#ff9800').setFontColor('white');
    
    // Meal headers
    const mealHeaders = [
      'Date', 'Day', 'Breakfast', 'Lunch', 'Dinner', 'Adults', 'Children', 'Total Guests', 'Summary'
    ];
    sheet.getRange(6, 1, 1, 9).setValues([mealHeaders]);
    
    // Meal data
    const mealRows = [];
    let totalBreakfast = 0;
    let totalLunch = 0;
    let totalDinner = 0;
    let totalAdults = 0;
    let totalChildren = 0;
    let totalGuests = 0;
    
    outlookData.forEach(dayData => {
      const dayOfWeek = dayData.date.toLocaleDateString('en-US', { weekday: 'short' });
      const dateStr = formatDate(dayData.date);
      const summary = `B=${dayData.breakfast}, L=${dayData.lunch}, D=${dayData.dinner}`;
      
      totalBreakfast += dayData.breakfast;
      totalLunch += dayData.lunch;
      totalDinner += dayData.dinner;
      totalAdults += dayData.adults;
      totalChildren += dayData.children;
      totalGuests += dayData.totalOccupancy;
      
      mealRows.push([
        dateStr,
        dayOfWeek,
        dayData.breakfast,
        dayData.lunch,
        dayData.dinner,
        dayData.adults,
        dayData.children,
        dayData.totalOccupancy,
        summary
      ]);
    });
    
    if (mealRows.length > 0) {
      sheet.getRange(7, 1, mealRows.length, 9).setValues(mealRows);
    }
    
    // Meal totals
    const mealTotalRow = 7 + mealRows.length + 1;
    sheet.getRange(mealTotalRow, 1, 1, 9).setValues([[
      'TOTAL', '', totalBreakfast, totalLunch, totalDinner, totalAdults, totalChildren, totalGuests, `Total Period: ${totalBreakfast + totalLunch + totalDinner} meals`
    ]]);
    
    // Formato para la hoja Meal Outlook
    formatMealOutlookSheet(sheet, mealRows.length, mealTotalRow);
  }

  
// VERSIÓN FINAL: Genera el reporte combinado de Ocupación y Housekeeping
function generateOccupancyOutlookInSheet(sheet, outlookData, housekeepingData, dateRange, periodType) {
    sheet.clear();

    if (!dateRange || dateRange.length === 0) {
        sheet.getRange(1, 1).setValue('Error: A valid date range was not provided.');
        return;
    }

    const overallStartDate = formatReadableDate(dateRange[0]);
    const overallEndDate = formatReadableDate(dateRange[dateRange.length - 1]);
    
    sheet.getRange(1, 1).setValue(`DAILY OPERATIONS OUTLOOK`);
    sheet.getRange(1, 1, 1, (dateRange.length * 7)).merge().setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold').setBackground('#4a148c').setFontColor('white');

    sheet.getRange(2, 1).setValue(`Period: ${overallStartDate} to ${overallEndDate}`);
    sheet.getRange(2, 1, 1, (dateRange.length * 7)).merge().setHorizontalAlignment('center').setFontSize(12).setFontWeight('bold');

    let currentColumn = 1;
    const dailyTableStartRow = 4;

    outlookData.forEach(dayData => {
        // Usar la fecha del objeto 'dayData' es más seguro y robusto
        const date = dayData.date; 
        const dayOfWeek = date.toLocaleDateString('en-US', { weekday: 'long' });
        const dateStr = formatDate(date);
        let currentRow = dailyTableStartRow;

        // --- TÍTULO DEL DÍA ---
        sheet.getRange(currentRow, currentColumn, 1, 5).merge().setValue(`${dayOfWeek} ${dateStr}`).setFontSize(14).setFontWeight('bold').setBackground('#673ab7').setFontColor('white').setHorizontalAlignment('center');
        currentRow++;

        // --- 1. TABLA DE OCUPACIÓN ---
        const occupancyHeaders = ['Cabin', 'Cabin Name', 'Guests', 'Arrival Date', 'Departure Date'];
        sheet.getRange(currentRow, currentColumn, 1, 5).setValues([occupancyHeaders]).setFontWeight('bold').setBackground('#b39ddb').setFontColor('black');
        currentRow++;

        const occupancyRows = [];
        let dailyTotalGuests = 0;
        if (dayData.cabinBreakdown && dayData.cabinBreakdown.length > 0) {
            dayData.cabinBreakdown.forEach(cabin => {
                occupancyRows.push([cabin.cabinNumber, cabin.cabinName, cabin.total, formatDate(cabin.arrivalDate), formatDate(cabin.departureDate)]);
                dailyTotalGuests += cabin.total;
            });
        } else {
            occupancyRows.push(['No Occupancy', '', 0, '', '']);
        }
        sheet.getRange(currentRow, currentColumn, occupancyRows.length, 5).setValues(occupancyRows);
        sheet.getRange(currentRow, currentColumn, occupancyRows.length, 5).setHorizontalAlignment('center');
        currentRow += occupancyRows.length;
        
        // --- TOTAL DE OCUPACIÓN ---
        sheet.getRange(currentRow, currentColumn, 1, 2).merge().setValue('Total Guests:').setHorizontalAlignment('right').setFontWeight('bold');
        sheet.getRange(currentRow, currentColumn + 2).setValue(dailyTotalGuests).setHorizontalAlignment('center').setFontWeight('bold');
        currentRow += 2;

        // --- 2. TABLA DE HOUSEKEEPING ---
        const dateISO = date.toISOString().slice(0, 10);
        const housekeepingTasks = housekeepingData.get(dateISO) || [];

        sheet.getRange(currentRow, currentColumn, 1, 5).merge().setValue('Housekeeping Tasks').setFontWeight('bold').setBackground('#00838f').setFontColor('white').setHorizontalAlignment('center');
        currentRow++;
        
        if (housekeepingTasks.length > 0) {
            const hkHeaders = ['Time', 'Cabin Name', 'Task'];
            
            // ===== CORRECCIÓN AQUÍ =====
            // El rango ahora es de 3 columnas para que coincida con los 3 encabezados de hkHeaders.
            sheet.getRange(currentRow, currentColumn, 1, 3).setValues([hkHeaders]).setFontWeight('bold').setBackground('#4dd0e1');
            currentRow++;

            const hkRows = housekeepingTasks.map(t => [t.time, t.cabinName, t.task]);
            sheet.getRange(currentRow, currentColumn, hkRows.length, 3).setValues(hkRows);
            sheet.getRange(currentRow, currentColumn + 2, hkRows.length, 1).setWrap(true);
            currentRow += hkRows.length;
        } else {
            sheet.getRange(currentRow, currentColumn, 1, 5).merge().setValue('No activities assigned').setFontStyle('italic').setHorizontalAlignment('center');
            currentRow++;
        }
        
        const blockEndRow = currentRow - 1;
        sheet.getRange(dailyTableStartRow, currentColumn, blockEndRow - dailyTableStartRow + 1, 5).setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

        currentColumn += 7;
    });

    formatOccupancyOutlookSheet(sheet, (dateRange.length * 7));
}


  // NUEVA FUNCIÓN: Formato para la hoja Meal Outlook
  function formatMealOutlookSheet(sheet, mealRowCount, mealTotalRow) {
    const mealHeaderRange = sheet.getRange(6, 1, 1, 9);
    mealHeaderRange.setFontWeight('bold');
    mealHeaderRange.setBackground('#ff9800');
    mealHeaderRange.setFontColor('white');
    
    if (mealRowCount > 0) {
      const mealDataRange = sheet.getRange(6, 1, mealRowCount + 1, 9);
      mealDataRange.setBorder(true, true, true, true, true, true);
      
      sheet.getRange(7, 3, mealRowCount, 6).setHorizontalAlignment('center');
      sheet.getRange(7, 3, mealRowCount, 6).setNumberFormat('0');
      
      for (let i = 0; i < mealRowCount; i++) {
        if (i % 2 === 1) {
          sheet.getRange(7 + i, 1, 1, 9).setBackground('#fff3e0');
        }
      }
    }
    
    const mealTotalRange = sheet.getRange(mealTotalRow, 1, 1, 9);
    mealTotalRange.setFontWeight('bold');
    mealTotalRange.setBackground('#f57c00');
    mealTotalRange.setFontColor('white');
    mealTotalRange.setBorder(true, true, true, true, true, true);
    mealTotalRange.setHorizontalAlignment('center');
    
    sheet.autoResizeColumns(1, 9);
    sheet.setColumnWidth(9, 200);
  }

  // NUEVA FUNCIÓN: Formato para la hoja Occupancy Outlook (ADAPTADO AL NUEVO DISEÑO)
  // NUEVA FUNCIÓN: Formato para la hoja Occupancy Outlook (ADAPTADO AL NUEVO DISEÑO HORIZONTAL)
  // FORMATO PARA LA HOJA OCCUPANCY OUTLOOK (CON ANCHOS DE COLUMNA AJUSTABLES)
function formatOccupancyOutlookSheet(sheet, maxColsUsedForTitles) {
    // Este bucle recorre cada tabla diaria que se crea en el reporte
    // y aplica los mismos anchos de columna a cada una.
    // El incremento `col += 7` es porque cada tabla usa 5 columnas + 2 de espacio.
    for (let col = 1; col <= maxColsUsedForTitles; col += 7) { 

        // --- AQUÍ ES DONDE PUEDES CAMBIAR LOS TAMAÑOS ---
        // El número es el ancho en píxeles. Experimenta con ellos.

        // Columna 1: 'Cabin' (La hacemos más angosta)
        sheet.setColumnWidth(col, 65);

        // Columna 2: 'Cabin Name' (La hacemos más ancha)
        sheet.setColumnWidth(col + 1, 180);

        // Columna 3: 'Guests' (Más angosta)
        sheet.setColumnWidth(col + 2, 65);

        // Columna 4: 'Arrival Date' (Tamaño estándar para fecha)
        sheet.setColumnWidth(col + 3, 100);

        // Columna 5: 'Departure Date' (Tamaño estándar para fecha)
        sheet.setColumnWidth(col + 4, 100);

        // Columnas 6 y 7: Estos son los espacios en blanco entre las tablas.
        // Es mejor dejarlos pequeños.
        sheet.setColumnWidth(col + 5, 25);
        sheet.setColumnWidth(col + 6, 25);
    }
    
    // Esto alinea todo el texto en la parte superior de las celdas, se ve mejor.
    sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).setVerticalAlignment('top');
}

  // ================================
  // GENERAL HELPER FUNCTIONS
  // ================================

 
  // FUNCTION: Get active guests data for a specific date (reused by multiple reports)
function getDailyData(dataSheet, queryDate) {
    const data = dataSheet.getDataRange().getValues();
    const headers = data[0];
    
    const cleanDate = new Date(queryDate.getFullYear(), queryDate.getMonth(), queryDate.getDate());
    
    // Load cabin names
    const cabinNames = getCabinNames();
    
    // Find columns
    const colJOB = headers.indexOf('JOB ID');
    const colFullName = headers.indexOf('FullName');
    const colArrivalDate = headers.indexOf('ArrivalDate');
    const colDepartDate = headers.indexOf('DepartDate');
    const colArrivalMeal = headers.indexOf('ArrivalMeal');
    const colDepartMeal = headers.indexOf('DepartMeal');
    const colItemID = headers.indexOf('Item ID');
    // ===== NUEVA LÍNEA: Busca la columna CustomerID =====
    // Si tu columna se llama diferente, cámbialo aquí.
    const colCustomerID = headers.indexOf('CustomerID'); 

    if (colJOB === -1 || colFullName === -1 || colArrivalDate === -1 || colDepartDate === -1 || colItemID === -1 || colCustomerID === -1) {
        throw new Error('Could not find all required columns in "Original Data". Make sure JOB ID, FullName, ArrivalDate, DepartDate, Item ID, and CustomerID exist.');
    }
    
    const guests = [];
    const families = {};
    const duplicatesDetected = new Set();
    
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const jobId = row[colJOB];
        const fullName = row[colFullName];
        const arrivalDateRaw = row[colArrivalDate];
        const departDateRaw = row[colDepartDate];
        const arrivalMeal = row[colArrivalMeal];
        const departMeal = row[colDepartMeal];
        const itemID = row[colItemID];
        // ===== NUEVA LÍNEA: Obtiene el valor de CustomerID =====
        const customerID = row[colCustomerID];

        if (!jobId || !fullName || !arrivalDateRaw || !departDateRaw) continue;
        
        if (!itemID || !itemID.toString().trim().toUpperCase().startsWith('FB')) {
            continue;
        }
        
        const uniqueKey = `${jobId}|${fullName}|${arrivalDateRaw}|${departDateRaw}|${arrivalMeal}|${departMeal}`;
        
        if (duplicatesDetected.has(uniqueKey)) {
            continue;
        }
        duplicatesDetected.add(uniqueKey);
        
        let arrivalDate, departureDate;
        
        if (typeof arrivalDateRaw === 'string') {
            arrivalDate = parseAmericanDate(arrivalDateRaw);
        } else {
            arrivalDate = new Date(arrivalDateRaw);
        }
        
        if (typeof departDateRaw === 'string') {
            departureDate = parseAmericanDate(departDateRaw);
        } else {
            departureDate = new Date(departDateRaw);
        }
        
        if (!arrivalDate || !departureDate || isNaN(arrivalDate.getTime()) || isNaN(departureDate.getTime())) {
            continue;
        }
        
        const cleanArrival = new Date(arrivalDate.getFullYear(), arrivalDate.getMonth(), arrivalDate.getDate());
        const cleanDeparture = new Date(departureDate.getFullYear(), departureDate.getMonth(), departureDate.getDate());
        
        const isInHotel = cleanArrival <= cleanDate && cleanDeparture >= cleanDate;
        
        if (isInHotel) {
            const jobIdStr = jobId.toString().trim();
            const cabinName = cabinNames[jobIdStr] || `Cabin ${jobIdStr}`;
            const guestType = getGuestType(itemID);
            
            families[jobIdStr] = cabinName;
            
            // ===== NUEVA LÍNEA: Se añade customerID al objeto del huésped =====
            guests.push({
                cabinNumber: jobId,
                guestName: fullName,
                checkIn: arrivalDate,
                checkOut: departureDate,
                arrivalMeal: arrivalMeal || 'No Meal',
                departMeal: departMeal || 'No Meal',
                cabinName: cabinName,
                guestType: guestType,
                itemID: itemID,
                customerID: customerID // <-- Se añade aquí
            });
        }
    }
    
    return { guests, families };
}

  // FUNCTION: Get outlook data (reused by combined and separated outlook reports)
  // FUNCTION: Get outlook data (reused by combined and separated outlook reports)
function getOutlookData(dataSheet, dateRange) {
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  
  const cabinNames = getCabinNames();
  
  const colJOB = headers.indexOf('JOB ID');
  const colFullName = headers.indexOf('FullName');
  const colArrivalDate = headers.indexOf('ArrivalDate');
  const colDepartDate = headers.indexOf('DepartDate');
  const colArrivalMeal = headers.indexOf('ArrivalMeal');
  const colDepartMeal = headers.indexOf('DepartMeal');
  const colItemID = headers.indexOf('Item ID');
  
  if (colJOB === -1 || colFullName === -1 || colArrivalDate === -1 || colDepartDate === -1) {
    throw new Error('Could not find all required columns');
  }
  
  if (colItemID === -1) {
    throw new Error('Could not find "Item ID" column. Please verify the column exists.');
  }
  
  const outlookData = [];
  
  dateRange.forEach(queryDate => {
    const cleanDate = new Date(queryDate.getFullYear(), queryDate.getMonth(), queryDate.getDate());
    
    const dailyGuests = [];
    const duplicatesDetected = new Set();
    const cabinBreakdown = new Map();
    
    // Mapa para almacenar las fechas de llegada y salida por reserva (jobId)
    const reservationDates = new Map();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const jobId = row[colJOB];
      const fullName = row[colFullName];
      const arrivalDateRaw = row[colArrivalDate];
      const departDateRaw = row[colDepartDate];
      const arrivalMeal = row[colArrivalMeal];
      const departMeal = row[colDepartMeal];
      const itemID = row[colItemID];
      
      if (!jobId || !fullName || !arrivalDateRaw || !departDateRaw) continue;
      
      if (!itemID || !itemID.toString().trim().toUpperCase().startsWith('FB')) {
        continue;
      }
      
      const uniqueKey = `${jobId}|${fullName}|${arrivalDateRaw}|${departDateRaw}|${arrivalMeal}|${departMeal}`;
      
      if (duplicatesDetected.has(uniqueKey)) {
        continue;
      }
      duplicatesDetected.add(uniqueKey);
      
      let arrivalDate, departureDate;
      
      if (typeof arrivalDateRaw === 'string') {
        arrivalDate = parseAmericanDate(arrivalDateRaw);
      } else {
        arrivalDate = new Date(arrivalDateRaw);
      }
      
      if (typeof departDateRaw === 'string') {
        departureDate = parseAmericanDate(departDateRaw);
      } else {
        departureDate = new Date(departDateRaw);
      }
      
      if (!arrivalDate || !departureDate || isNaN(arrivalDate.getTime()) || isNaN(departureDate.getTime())) {
        continue;
      }
      
      const cleanArrival = new Date(arrivalDate.getFullYear(), arrivalDate.getMonth(), arrivalDate.getDate());
      const cleanDeparture = new Date(departureDate.getFullYear(), departureDate.getMonth(), departureDate.getDate());
      
      const isInHotel = cleanArrival <= cleanDate && cleanDeparture >= cleanDate;
      
      if (isInHotel) {
        const jobIdStr = jobId.toString().trim();
        const cabinName = cabinNames[jobIdStr] || `Cabin ${jobIdStr}`;
        const guestType = getGuestType(itemID);
        
        // Guardar las fechas de la reserva si es la primera vez que la vemos
        if (!reservationDates.has(jobIdStr)) {
            reservationDates.set(jobIdStr, { arrival: arrivalDate, departure: departureDate });
        }

        if (!cabinBreakdown.has(jobIdStr)) {
          cabinBreakdown.set(jobIdStr, {
            cabinNumber: jobId,
            cabinName: cabinName,
            adults: 0,
            children: 0,
            total: 0,
            // Añadir las fechas a la información de la cabaña
            arrivalDate: arrivalDate,
            departureDate: departureDate
          });
        }
        
        const cabin = cabinBreakdown.get(jobIdStr);
        if (guestType === 'Adult') {
          cabin.adults++;
        } else if (guestType === 'Child') {
          cabin.children++;
        }
        cabin.total++;
        
        dailyGuests.push({
          cabinNumber: jobId,
          guestName: fullName,
          checkIn: arrivalDate,
          checkOut: departureDate,
          arrivalMeal: arrivalMeal || 'No Meal',
          departMeal: departMeal || 'No Meal',
          cabinName: cabinName,
          guestType: guestType,
          itemID: itemID
        });
      }
    }
    
    let breakfast = 0, lunch = 0, dinner = 0, adults = 0, children = 0, totalOccupancy = 0;
    
    dailyGuests.forEach(guest => {
      const dailyMeals = calculateDailyMeals(guest, queryDate);
      breakfast += dailyMeals.breakfast;
      lunch += dailyMeals.lunch;
      dinner += dailyMeals.dinner;
      if (guest.guestType === 'Adult') adults++;
      else if (guest.guestType === 'Child') children++;
      totalOccupancy++;
    });
    
    const cabinBreakdownArray = Array.from(cabinBreakdown.values()).sort((a, b) =>
      a.cabinNumber.localeCompare(b.cabinNumber)
    );
    
    outlookData.push({
      date: queryDate,
      breakfast, lunch, dinner, adults, children, totalOccupancy,
      guestCount: dailyGuests.length,
      cabinBreakdown: cabinBreakdownArray
    });
  });
  
  return outlookData;
}

  // FUNCTION: Get cabin names from reference sheet
  function getCabinNames() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const cabinSheet = ss.getSheetByName('Cabin Names');
      
      if (!cabinSheet) {
        return {};
      }
      
      const data = cabinSheet.getDataRange().getValues();
      const cabinNames = {};
      
      for (let i = 1; i < data.length; i++) {
        const jobId = data[i][0];
        const cabinName = data[i][1];
        
        if (jobId && cabinName) {
          cabinNames[jobId.toString().trim()] = cabinName.toString().trim();
        }
      }
      
      return cabinNames;
      
    } catch (error) {
      debugLog('Error loading cabin names:', error.message);
      return {};
    }
  }

  // FUNCTION: Determine if guest is Adult or Child based on Item ID
  function getGuestType(itemID) {
    if (!itemID) return 'Unknown';
    
    const itemStr = itemID.toString().trim().toUpperCase();
    
    if (itemStr.includes('_A_') || itemStr.includes('A_MEAL')) {
      return 'Adult';
    } else if (itemStr.includes('_C_') || itemStr.includes('C_MEAL')) {
      return 'Child';
    }
    
    return 'Adult'; // Default to Adult if pattern not clear
  }

  // FUNCTION: Parse American date format (MM/DD/YYYY)
  function parseAmericanDate(dateString) {
    if (!dateString) return null;
    
    const cleanString = dateString.toString().trim();
    const parts = cleanString.split('/');
    
    if (parts.length !== 3) return null;
    
    const month = parseInt(parts[0]) - 1;
    const day = parseInt(parts[1]);
    const year = parseInt(parts[2]);
    
    if (isNaN(month) || isNaN(day) || isNaN(year)) return null;
    
    return new Date(year, month, day);
  }

  // FUNCTION: Calculate meals for a guest on a specific day (No Meal logic included)
  function calculateDailyMeals(guest, queryDate) {
    const arrivalDate = new Date(guest.checkIn.getFullYear(), guest.checkIn.getMonth(), guest.checkIn.getDate());
    const departureDate = new Date(guest.checkOut.getFullYear(), guest.checkOut.getMonth(), guest.checkOut.getDate());
    const normalizedQueryDate = new Date(queryDate.getFullYear(), queryDate.getMonth(), queryDate.getDate());
    
    const isFirstDay = normalizedQueryDate.getTime() === arrivalDate.getTime();
    const isLastDay = normalizedQueryDate.getTime() === departureDate.getTime();
    
    let breakfast = 1;
    let lunch = 1;
    let dinner = 1;
    
    // First day logic
    if (isFirstDay) {
      const arrivalMeal = guest.arrivalMeal ? guest.arrivalMeal.toLowerCase().trim() : '';
      
      if (arrivalMeal.includes('no meal')) {
        breakfast = 0;
        lunch = 0;
        dinner = 0;
      } else if (arrivalMeal.includes('lunch')) {
        breakfast = 0;
        lunch = 1;
        dinner = 1;
      } else if (arrivalMeal.includes('dinner')) {
        breakfast = 0;
        lunch = 0;
        dinner = 1;
      } else if (arrivalMeal.includes('breakfast')) {
        breakfast = 1;
        lunch = 1;
        dinner = 1;
      } else if (arrivalMeal === '') {
        breakfast = 1;
        lunch = 1;
        dinner = 1;
      } else {
        breakfast = 1;
        lunch = 1;
        dinner = 1;
      }
    }
    
    // Last day logic
    if (isLastDay) {
      const departMeal = guest.departMeal ? guest.departMeal.toLowerCase().trim() : '';
      
      if (departMeal.includes('no meal')) {
        // no change, means full day before departure (or departure after all meals)
      } else if (departMeal.includes('breakfast')) {
        lunch = 0;
        dinner = 0;
      } else if (departMeal.includes('lunch')) {
        dinner = 0;
      } else if (departMeal.includes('dinner')) {
        // no change, means full day before departure
      } else if (departMeal === '') {
        // no change, assume full day
      }
    }
    
    return { breakfast, lunch, dinner };
  }

  // FORMATTING HELPERS
  function alternateCabinColors(sheet, numRows, startRow) {
    let currentCabin = null;
    let colorToggle = false;
    
    for (let i = 0; i < numRows; i++) {
      const cabinValue = sheet.getRange(startRow + i, 1).getValue();
      
      if (cabinValue !== currentCabin) {
        currentCabin = cabinValue;
        colorToggle = !colorToggle;
      }
      
      if (colorToggle) {
        sheet.getRange(startRow + i, 1, 1, 9).setBackground('#f8f9fa');
      }
    }
  }

  function formatDate(date) {
    if (!date) return '';
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    const year = date.getFullYear();
    return `${month}/${day}/${year}`;
  }

  function formatReadableDate(date) {
    const options = { 
      weekday: 'long', 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric' 
    };
    return date.toLocaleDateString('en-US', options);
  }

  function debugLog(message, data = null) {
    if (data) {
      console.log(message, data);
    } else {
      console.log(message);
    }
  }

  // FUNCTION: Cleanup Old Reports (run once if needed)
  function cleanupOldReports() {
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.alert(
      'Cleanup Old Reports',
      'This will delete all old report sheets (Report_YYYY-MM-DD format).\n\nAre you sure you want to continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = ss.getSheets();
      let deletedCount = 0;
      
      sheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (sheetName.match(/^Report_\d{4}-\d{2}-\d{2}$/)) {
          ss.deleteSheet(sheet);
          deletedCount++;
        }
      });
      
      ui.alert(
        'Cleanup Complete',
        `Successfully deleted ${deletedCount} old report sheets.\n\nFrom now on, all reports will use the single "Meal Report" sheet.`,
        ui.ButtonSet.OK
      );
    }
  }

  // ================================
  // AIRPORT TRANSFER FUNCTIONS
  // ================================

  // MAIN FUNCTION: Generate all upcoming transfers (from today onwards)
  function generateUpcomingTransfers() {
    const ui = SpreadsheetApp.getUi();
    
    try {
      ui.alert(
        'Processing...',
        'Generating all upcoming airport transfers from today onwards...',
        ui.ButtonSet.OK
      );
      
      processUpcomingTransfers();
      
      ui.alert(
        'Transfer Schedule Generated Successfully',
        'All upcoming airport transfers have been generated successfully.\n\nYou can find the complete schedule in the "Airport Transfers" sheet.\n\nThis shows all transfers from today onwards, sorted by date and time.',
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      ui.alert(
        'Error',
        `Could not generate the transfer schedule:\n\n${error.message}\n\nPlease verify that the "Original Data" sheet exists with valid transfer information.`,
        ui.ButtonSet.OK
      );
      debugLog('Error in upcoming transfers:', error.message);
    }
  }

  // FUNCTION: Process all upcoming transfers
  function processUpcomingTransfers() {
    const today = new Date();
    const todayClean = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const originalData = ss.getSheetByName('Original Data');
    if (!originalData) {
      throw new Error('Could not find "Original Data" sheet');
    }
    
    const upcomingTransfers = getAllUpcomingTransfers(originalData, todayClean);
    
    let transferSheet = ss.getSheetByName('Airport Transfers');
    if (!transferSheet) {
      transferSheet = ss.insertSheet('Airport Transfers');
    }
    
    generateUpcomingTransfersReport(transferSheet, upcomingTransfers, todayClean);
  }

  // FUNCTION: Get all upcoming transfers from today onwards
  // FUNCTION: Get all upcoming transfers from today onwards (FINAL ROBUST VERSION)
function getAllUpcomingTransfers(dataSheet, fromDate) {
    const data = dataSheet.getDataRange().getValues();
    const headers = data[0];
    
    const cabinNames = getCabinNames();
    
    const colJOB = headers.indexOf('JOB ID');
    const colFullName = headers.indexOf('FullName');
    const colArrivalDate = headers.indexOf('ArrivalDate');
    const colDepartDate = headers.indexOf('DepartDate');
    const colItemID = headers.indexOf('Item ID');
    let colDescription = headers.indexOf('Description');
    if (colDescription === -1) colDescription = headers.indexOf('Desciption');
    
    if ([colJOB, colFullName, colArrivalDate, colDepartDate, colItemID, colDescription].includes(-1)) {
        throw new Error('Could not find all required columns for transfers. Check for JOB ID, FullName, ArrivalDate, DepartDate, Item ID, and Description.');
    }
    
    const allTransfers = [];
    
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const jobId = row[colJOB];
        const fullName = row[colFullName];
        const itemID = row[colItemID];
        const description = row[colDescription] || '';
        
        if (!jobId || !fullName || !itemID || !itemID.toString().trim().toUpperCase().includes('ADMIN_AUTO')) {
            continue;
        }
        
        // Process ARRIVAL
        const arrivalDateRaw = row[colArrivalDate];
        if (arrivalDateRaw) {
            const arrivalDate = parseAmericanDate(arrivalDateRaw.toString());
            if (arrivalDate && new Date(arrivalDate.getFullYear(), arrivalDate.getMonth(), arrivalDate.getDate()) >= fromDate) {
                const flightInfo = parseFlightInfo(description, 'ARRIVAL');
                if (flightInfo) {
                    allTransfers.push({
                        cabinNumber: jobId,
                        cabinName: cabinNames[jobId.toString().trim()] || `Cabin ${jobId}`,
                        guestName: fullName,
                        transferType: 'ARRIVAL',
                        transferDate: arrivalDate,
                        flightInfo: flightInfo,
                        sortDate: arrivalDate,
                        sortTime: flightInfo.timeForSort || '00:00'
                    });
                }
            }
        }
        
        // Process DEPARTURE
        const departDateRaw = row[colDepartDate];
        if (departDateRaw) {
            const departureDate = parseAmericanDate(departDateRaw.toString());
            if (departureDate && new Date(departureDate.getFullYear(), departureDate.getMonth(), departureDate.getDate()) >= fromDate) {
                const flightInfo = parseFlightInfo(description, 'DEPARTURE');
                if (flightInfo) {
                    allTransfers.push({
                        cabinNumber: jobId,
                        cabinName: cabinNames[jobId.toString().trim()] || `Cabin ${jobId}`,
                        guestName: fullName,
                        transferType: 'DEPARTURE',
                        transferDate: departureDate,
                        flightInfo: flightInfo,
                        sortDate: departureDate,
                        sortTime: flightInfo.timeForSort || '23:59'
                    });
                }
            }
        }
    }
    
    allTransfers.sort((a, b) => {
        const dateA = a.sortDate.getTime();
        const dateB = b.sortDate.getTime();
        if (dateA !== dateB) {
            return dateA - dateB;
        }
        return a.sortTime.localeCompare(b.sortTime);
    });
    
    return allTransfers;
}

  // FUNCTION: Parse flight information with sort time
  // FUNCTION: Parse flight information with context (Arrival/Departure) - FINAL, SIMPLIFIED LOGIC
function parseFlightInfo(description, transferType) {
    if (!description || typeof description !== 'string') {
        return null;
    }

    let searchString = description;
    const lowerDesc = description.toLowerCase();
    
    const departureKeywords = ['departing', 'departure', 'departs'];
    const arrivalKeywords = ['arrive', 'arriving'];

    let separatorIndex = -1;
    let keywordFound = '';

    // Find the first departure keyword
    for (const keyword of departureKeywords) {
        const index = lowerDesc.indexOf(keyword);
        if (index !== -1) {
            separatorIndex = index;
            keywordFound = 'departure';
            break;
        }
    }

    // If no departure keyword, find the first arrival keyword
    if (separatorIndex === -1) {
        for (const keyword of arrivalKeywords) {
            const index = lowerDesc.indexOf(keyword);
            if (index !== -1) {
                separatorIndex = index;
                keywordFound = 'arrival';
                break;
            }
        }
    }

    // Logic based on what was found
    if (keywordFound === 'departure') {
        if (transferType === 'ARRIVAL') {
            searchString = description.substring(0, separatorIndex);
        } else { // DEPARTURE
            searchString = description.substring(separatorIndex);
        }
    } else if (keywordFound === 'arrival') {
        if (transferType === 'DEPARTURE') {
            // This text only contains arrival info, but we are looking for departure
            return null; 
        }
        // Otherwise, the whole string is for arrival, which is the default
    }
    // If no keywords are found, searchString remains the full description, which is the desired fallback.
    
    if (!searchString || !searchString.trim()) return null;

    const flightPattern = /([A-Z]{2,3})\s*(\d{2,4})/;
    const timePattern = /(\d{1,2}:\d{2}\s*(?:am|pm)?)/i;

    const flightMatch = searchString.match(flightPattern);
    const timeMatch = searchString.match(timePattern);

    if (!flightMatch && !timeMatch) {
       return null;
    }

    const airline = flightMatch ? flightMatch[1].toUpperCase() : 'N/A';
    const flightNumber = flightMatch ? flightMatch[2] : 'N/A';
    const time = timeMatch ? timeMatch[0].trim() : 'N/A';
    
    return {
        airline: airline,
        flightNumber: flightNumber,
        time: time,
        timeForSort: convertToSortableTime(time),
        details: searchString.trim().replace(/,$/, '')
    };
}

  // HELPER FUNCTION: Convert time to sortable format
  function convertToSortableTime(timeStr) {
    if (!timeStr) return '12:00';
    
    if (/^\d{4}$/.test(timeStr)) {
      const hours = timeStr.substring(0, 2);
      const minutes = timeStr.substring(2, 4);
      return `${hours}:${minutes}`;
    }
    
    if (timeStr.includes('AM') || timeStr.includes('PM')) {
      const isPM = timeStr.includes('PM');
      const timeOnly = timeStr.replace(/(AM|PM)/gi, '').trim();
      const [hours, minutes] = timeOnly.split(':');
      let hour24 = parseInt(hours);
      
      if (isPM && hour24 !== 12) hour24 += 12;
      if (!isPM && hour24 === 12) hour24 = 0;
      
      return `${hour24.toString().padStart(2, '0')}:${minutes || '00'}`;
    }
    
    return timeStr;
  }

  // FUNCTION: Generate upcoming transfers report
  function generateUpcomingTransfersReport(sheet, transfers, fromDate) {
    sheet.clear();
    
    const today = formatReadableDate(fromDate);
    sheet.getRange(1, 1).setValue(`UPCOMING AIRPORT TRANSFERS - From ${today} Onwards`);
    sheet.getRange(1, 1, 1, 10).merge().setHorizontalAlignment('center');
    sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setBackground('#ff6b6b').setFontColor('white');
    
    if (transfers.length === 0) {
      sheet.getRange(3, 1).setValue('No upcoming airport transfers found.');
      sheet.getRange(3, 1).setFontSize(14).setFontWeight('bold');
      return;
    }
    
    const arrivalCount = transfers.filter(t => t.transferType === 'ARRIVAL').length;
    const departureCount = transfers.filter(t => t.transferType === 'DEPARTURE').length;
    
    sheet.getRange(2, 1).setValue(`Total Transfers: ${transfers.length} | Arrivals: ${arrivalCount} | Departures: ${departureCount}`);
    sheet.getRange(2, 1, 1, 10).merge().setHorizontalAlignment('center');
    sheet.getRange(2, 1).setFontSize(12).setFontWeight('bold');
    
    const headers = [
      'Date', 'Day', 'Type', 'Time', 'Cabin', 'Cabin Name', 'Guest Name', 'Airline', 'Flight #', 'Full Flight Details'
    ];
    sheet.getRange(4, 1, 1, 10).setValues([headers]);
    
    const detailRows = [];
    let currentDate = null;
    
    transfers.forEach(transfer => {
      const transferDateStr = formatDate(transfer.transferDate);
      const dayOfWeek = transfer.transferDate.toLocaleDateString('en-US', { weekday: 'short' });
      
      detailRows.push([
        transferDateStr,
        dayOfWeek,
        transfer.transferType,
        transfer.flightInfo.time,
        transfer.cabinNumber,
        transfer.cabinName, // Usar cabinName aquí
        transfer.guestName,
        transfer.flightInfo.airline,
        transfer.flightInfo.flightNumber,
        transfer.flightInfo.details
      ]);
    });
    
    if (detailRows.length > 0) {
      sheet.getRange(5, 1, detailRows.length, 10).setValues(detailRows);
    }
    
    formatUpcomingTransfersSheet(sheet, detailRows.length);
  }

  // FUNCTION: Format upcoming transfers sheet
  function formatUpcomingTransfersSheet(sheet, numRows) {
    const headerRange = sheet.getRange(4, 1, 1, 10);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#ff6b6b');
    headerRange.setFontColor('white');
    
    if (numRows > 0) {
      const detailRange = sheet.getRange(4, 1, numRows + 1, 10);
      detailRange.setBorder(true, true, true, true, true, true);
      
      sheet.getRange(5, 1, numRows, 2).setHorizontalAlignment('center');
      sheet.getRange(5, 3, numRows, 2).setHorizontalAlignment('center');
      sheet.getRange(5, 5, numRows, 1).setHorizontalAlignment('center');
      sheet.getRange(5, 8, numRows, 2).setHorizontalAlignment('center');
      
      let currentDate = null;
      let colorToggle = false;
      
      for (let i = 0; i < numRows; i++) {
        const rowDate = sheet.getRange(5 + i, 1).getValue();
        const transferType = sheet.getRange(5 + i, 3).getValue();
        
        if (rowDate !== currentDate) {
          currentDate = rowDate;
          colorToggle = !colorToggle;
        }
        
        let backgroundColor = '#ffffff';
        
        if (colorToggle) {
          backgroundColor = '#f8f9fa';
        }
        
        if (transferType === 'ARRIVAL') {
          backgroundColor = colorToggle ? '#e8f5e8' : '#d4edda';
        } else if (transferType === 'DEPARTURE') {
          backgroundColor = colorToggle ? '#fff2e8' : '#ffeaa7';
        }
        
        sheet.getRange(5 + i, 1, 1, 10).setBackground(backgroundColor);
      }
    }
    
    sheet.autoResizeColumns(1, 10);
    sheet.setColumnWidth(6, 150);
    sheet.setColumnWidth(7, 150);
    sheet.setColumnWidth(10, 300);
  }



  // ================================
// AUTOMATIC DAILY REPORTS EMAIL SENDING
// ================================

function sendDailyReports() {
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1); // Generar reporte para el día siguiente
  const formattedDate = `${tomorrow.getFullYear()}-${(tomorrow.getMonth() + 1).toString().padStart(2, '0')}-${tomorrow.getDate().toString().padStart(2, '0')}`;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // === EMAIL AND SUBJECT CONFIGURATION ===
  const mealReportEmail = 'wausaukeekitchen@gmail.com'; // <--- CHANGE THIS to the email for meal reports
  const occupancyReportEmail = 'juanferojas19@gmail.com'; // <--- CHANGE THIS to the email for occupancy reports

  const mealSubject = `Informe de Comidas - ${formattedDate}`;
  const occupancySubject = `Informe de Ocupación - ${formattedDate}`;
  // ======================================================

  try {
    // 1. Generate reports (this will update 'Meal Report' and 'Occupancy' sheets)
    processMealReport(formattedDate);
    Logger.log(`Generated 'Meal Report' for ${formattedDate}.`);
    processOccupancyReport(formattedDate);
    Logger.log(`Generated 'Occupancy Report' for ${formattedDate}.`);

    // ===== SOLUCIÓN AQUÍ =====
    // Se añade esta línea para forzar que todos los cambios en las hojas
    // se apliquen ANTES de intentar convertirlas a PDF.
    SpreadsheetApp.flush();
    // ===========================

    // 2. Get sheets and verify existence
    const mealReportSheet = ss.getSheetByName('Meal Report');
    const occupancySheet = ss.getSheetByName('Occupancy');

    if (!mealReportSheet) {
      throw new Error('Could not find "Meal Report" sheet. Ensure it exists.');
    }
    if (!occupancySheet) {
      throw new Error('Could not find "Occupancy" sheet. Ensure it exists.');
    }

    // 3. Convert sheets to PDF
    const mealReportBlob = sheetToPdf(mealReportSheet);
    Logger.log('Meal Report converted to PDF.');
    const occupancyReportBlob = sheetToPdf(occupancySheet);
    Logger.log('Occupancy Report converted to PDF.');

    // 4. Send Meal Report email
    GmailApp.sendEmail(
      mealReportEmail,
      mealSubject,
      `Attached you will find the meal report for the day. ${formattedDate}.`,
      {
        attachments: [{
          fileName: `Meal Report - ${formattedDate}.pdf`,
          content: mealReportBlob.getBytes(),
          mimeType: 'application/pdf',
        }],
      }
    );
    Logger.log(`Food Report has sent a ${mealReportEmail} for ${formattedDate}.`);

    // 5. Send Occupancy Report email
    GmailApp.sendEmail(
      occupancyReportEmail,
      occupancySubject,
      `Attached you will find the occupancy report for the day ${formattedDate}.`,
      {
        attachments: [{
          fileName: `Occupancy Report - ${formattedDate}.pdf`,
          content: occupancyReportBlob.getBytes(),
          mimeType: 'application/pdf',
        }],
      }
    );
    Logger.log(`Occupancy Report sent to ${occupancyReportEmail} for ${formattedDate}.`);

  } catch (error) {
    Logger.log(`Error sending daily reports by email: ${error.message}`);
  }
}

  // Helper function to convert a sheet to PDF
  function sheetToPdf(sheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const url = ss.getUrl().replace(/edit$/, '');

    const url_ext = 'export?format=pdf&' +
      'size=A4&' +
      'portrait=true&' +
      'fitw=true&' +
      'sheetnames=false&' +
      'printtitle=false&' +
      'pagenumbers=false&' +
      'gridlines=false&' +
      'fzr=false' +
      '&gid=' + sheet.getSheetId();

    const params = {
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };

    const blob = UrlFetchApp.fetch(url + url_ext, params).getBlob().setName(`${sheet.getName()}.pdf`);
    return blob;
  }

  // Function to schedule the daily email sending (run this ONCE to create the trigger)
  function createDailyTrigger() {
    if (typeof sendDailyReports !== 'function') {
      Logger.log('Error: The sendDailyReports function is not defined.');
      return;
    }

    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'sendDailyReports') {
        ScriptApp.deleteTrigger(trigger);
        Logger.log('Existing trigger for sendDailyReports deleted.');
      }
    }

    // Create a new trigger to run daily at 1:00 PM (13:00 in 24-hour format)
    ScriptApp.newTrigger('sendDailyReports')
      .timeBased()
      .atHour(13) // This is 1 PM
      .everyDays(1)
      .create();

    Logger.log('Daily trigger for sendDailyReports created for 1:00 PM.');
  }
 
 

 // =============================================================
// AUTOMATIC WEEKLY OPERATIONS OUTLOOK EMAIL SENDING
// =============================================================

/**
 * Genera y envía el reporte "Operations Outlook" para los próximos 6 días.
 * Esta función es la que será llamada por el disparador automático.
 */
/**
 * Genera y envía el reporte "Operations Outlook" para los próximos 6 días.
 * Esta función es la que será llamada por el disparador automático.
 */
function sendOperationsOutlookReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- CONFIGURACIÓN ---
  const NUMBER_OF_DAYS = 6; // Hoy + 5 días = 6 días en total.
  const RECIPIENT_EMAIL = 'juanferojas19@gmail.com'; // <-- CAMBIA ESTO al correo deseado.
  // -------------------

  const today = new Date();
  const startDateText = `${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
  const subject = `Weekly Operations Outlook - Starting ${startDateText}`;
  const body = `Attached is the Operations Outlook report for the next ${NUMBER_OF_DAYS} days.`;

  try {
    // 1. Generar los datos para el reporte.
    Logger.log(`Generating Operations Outlook data for ${NUMBER_OF_DAYS} days.`);
    const dateRange = [];
    for (let i = 0; i < NUMBER_OF_DAYS; i++) {
      const currentDate = new Date(today);
      currentDate.setDate(today.getDate() + i);
      dateRange.push(currentDate);
    }
    const originalData = ss.getSheetByName('Original Data');
    if (!originalData) throw new Error('Could not find "Original Data" sheet');
    const outlookData = getOutlookData(originalData, dateRange);

    const housekeepingSheet = ss.getSheetByName('Housekeeping');
    const housekeepingData = housekeepingSheet ? getHousekeepingDataForRange(housekeepingSheet, dateRange) : new Map();

    // 2. Dibujar el reporte en la hoja "Operations Outlook"
    let operationsSheet = ss.getSheetByName('Operations Outlook');
    if (!operationsSheet) {
      operationsSheet = ss.insertSheet('Operations Outlook');
    }
    
    
    generateOccupancyOutlookInSheet(operationsSheet, outlookData, housekeepingData, dateRange, "Operations");
    // ===========================

    Logger.log('Operations Outlook sheet has been updated.');

    // 3. Forzar que la hoja se actualice antes de crear el PDF
    SpreadsheetApp.flush();

    // 4. Convertir la hoja actualizada a PDF
    const pdfBlob = sheetToPdf(operationsSheet);
    Logger.log('Operations Outlook sheet converted to PDF.');
    
    // 5. Enviar el correo electrónico
    GmailApp.sendEmail(RECIPIENT_EMAIL, subject, body, {
      attachments: [{
        fileName: `Operations Outlook - ${startDateText}.pdf`,
        content: pdfBlob.getBytes(),
        mimeType: 'application/pdf',
      }],
    });

    Logger.log(`Successfully sent Operations Outlook report to ${RECIPIENT_EMAIL}.`);

  } catch (error) {
    Logger.log(`Error sending the Operations Outlook report: ${error.message}`);
  }
}

/**
 * Crea el disparador (trigger) para enviar el reporte de Operations Outlook diariamente.
 * ¡EJECUTAR ESTA FUNCIÓN UNA SOLA VEZ DESDE EL EDITOR!
 */
/**
 * Crea el disparador (trigger) para enviar el reporte de Operations Outlook diariamente.
 * ¡EJECUTAR ESTA FUNCIÓN UNA SOLA VEZ DESDE EL EDITOR!
 */
function createOperationsOutlookTrigger() {
  const functionName = 'sendOperationsOutlookReport';
  
  // Borrar cualquier disparador antiguo para esta función para evitar duplicados
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`Deleted existing trigger for ${functionName}.`);
    }
  }

  // Crear un nuevo disparador para que se ejecute todos los días a las 8 AM
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .atHour(16) // <-- PUEDES CAMBIAR LA HORA AQUÍ (formato 24h, 8 = 8 AM)
    .everyDays(1)
    .create();

  const message = `Daily trigger created for ${functionName} to run at 4:00 AM.`;
  Logger.log(message);
  
  // CORRECCIÓN: Se muestra la alerta solo si el script se está ejecutando en un
  // contexto donde se puede mostrar una interfaz de usuario.
  try {
    SpreadsheetApp.getUi().alert('Success!', `The automatic daily email for the Operations Outlook report has been scheduled for 8:00 AM.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // Si no se puede mostrar la UI (ej. ejecutado por un trigger), simplemente se ignora el error.
    Logger.log('Cannot show UI alert in this context, but the trigger was created successfully.');
  }
}
  // ================================
// HOUSEKEEPING REPORT FUNCTIONS
// ================================

// MAIN FUNCTION: Generate housekeeping report with date picker
function generateHousekeepingReport() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Generate Housekeeping Report',
    'Enter the date for the report in YYYY-MM-DD format\n(Example: 2025-06-24):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const dateInput = response.getResponseText().trim();
    if (!dateInput) {
      ui.alert('Error', 'Please enter a valid date.', ui.ButtonSet.OK);
      return;
    }
    
    try {
      ui.alert('Processing...', `Generating housekeeping report for ${dateInput}...`, ui.ButtonSet.OK);
      processHousekeepingReport(dateInput);
      ui.alert('Report Generated Successfully', `The housekeeping report for ${dateInput} has been generated.\n\nYou can find it in the "Housekeeping Report" sheet.`, ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('Error', `Could not generate the report:\n\n${error.message}\n\nPlease verify:\n• Date format is YYYY-MM-DD\n• "Housekeeping" sheet exists with data.`, ui.ButtonSet.OK);
      debugLog('Error in housekeeping report:', error.message);
    }
  }
}

// QUICK FUNCTION: Today's housekeeping report
function todaysHousekeepingReport() {
  const today = new Date();
  const formattedDate = `${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert('Processing...', `Generating housekeeping report for today (${formattedDate})...`, ui.ButtonSet.OK);
    processHousekeepingReport(formattedDate);
    ui.alert('Report Generated Successfully', `Today's housekeeping report has been generated.\n\nYou can find it in the "Housekeeping Report" sheet.`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', `Could not generate the report:\n\n${error.message}`, ui.ButtonSet.OK);
    debugLog('Error in today\'s housekeeping report:', error.message);
  }
}

// PROCESSING FUNCTION: Handles the logic for daily housekeeping reports
function processHousekeepingReport(dateText) {
  const date = new Date(dateText + 'T00:00:00'); // Ensure correct date parsing
  if (isNaN(date.getTime())) {
    throw new Error('Invalid date format. Use YYYY-MM-DD');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const housekeepingSheet = ss.getSheetByName('Housekeeping');
  if (!housekeepingSheet) {
    throw new Error('Could not find "Housekeeping" sheet');
  }
  
  const tasks = getHousekeepingDataForDate(housekeepingSheet, date);
  
  let reportSheet = ss.getSheetByName('Housekeeping Report');
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet('Housekeeping Report');
  }
  
  generateDailyHousekeepingReportInSheet(reportSheet, tasks, date);
}

// DATA FUNCTION: Gets housekeeping tasks for a specific date
// DATA FUNCTION: Gets housekeeping tasks for a specific date
function getHousekeepingDataForDate(dataSheet, queryDate) {
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());
  const tasks = [];

  const colCabin = headers.indexOf('Cabin Names');
  const colDate = headers.indexOf('Date');
  const colTime = headers.indexOf('Time');
  const colTask = headers.indexOf('Task');
  
  if ([colCabin, colDate, colTime, colTask].includes(-1)) {
      throw new Error('Could not find all required columns in "Housekeeping" sheet: Cabin Names, Date, Time, Task.');
  }

  const queryDateString = queryDate.toLocaleDateString();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const taskDateValue = row[colDate];
    if (!taskDateValue) continue;

    // CORRECCIÓN: Usar la función parseAmericanDate que es más robusta para tu formato de hoja.
    const taskDate = (taskDateValue instanceof Date) ? taskDateValue : parseAmericanDate(taskDateValue.toString());
    if (!taskDate || isNaN(taskDate.getTime())) continue;

    if (taskDate.toLocaleDateString() === queryDateString) {
      tasks.push({
        cabinName: row[colCabin] || 'N/A',
        time: (row[colTime] instanceof Date) ? Utilities.formatDate(row[colTime], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
, 'HH:mm') : (row[colTime] || 'Any time'),
        task: row[colTask] || 'No description'
      });
    }
  }
  
  tasks.sort((a, b) => a.time.localeCompare(b.time));
  return tasks;
}

// SHEET GENERATION: Creates the daily housekeeping report sheet
function generateDailyHousekeepingReportInSheet(sheet, tasks, date) {
  sheet.clear();
  const formattedDate = formatReadableDate(date);
  
  sheet.getRange(1, 1).setValue(`DAILY HOUSEKEEPING REPORT - ${formattedDate}`);
  sheet.getRange(1, 1, 1, 3).merge().setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold').setBackground('#00acc1').setFontColor('white');
  
  if (tasks.length === 0) {
      sheet.getRange(3, 1).setValue('No housekeeping tasks scheduled for this date.');
      return;
  }
  
  const headers = ['Time', 'Cabin Name', 'Task to Perform'];
  sheet.getRange(3, 1, 1, 3).setValues([headers]);
  
  const detailRows = tasks.map(task => [task.time, task.cabinName, task.task]);
  sheet.getRange(4, 1, detailRows.length, 3).setValues(detailRows);
  
  formatDailyHousekeepingSheet(sheet, detailRows.length);
}

// FORMATTING: Styles the daily housekeeping report sheet
function formatDailyHousekeepingSheet(sheet, numRows) {
    const headerRange = sheet.getRange(3, 1, 1, 3);
    headerRange.setFontWeight('bold').setBackground('#00838f').setFontColor('white');

    if (numRows > 0) {
        const dataRange = sheet.getRange(3, 1, numRows + 1, 3);
        dataRange.setBorder(true, true, true, true, true, true);
        sheet.getRange(4, 1, numRows, 1).setHorizontalAlignment('center'); // Center time
        
        for (let i = 0; i < numRows; i++) {
            if (i % 2 === 1) {
                sheet.getRange(4 + i, 1, 1, 3).setBackground('#e0f7fa');
            }
        }
    }
    sheet.autoResizeColumns(1, 3);
    sheet.setColumnWidth(3, 400); // Give more space for the task description
}


// --- Housekeeping Outlook Functions ---


// MAIN FUNCTION: Generate housekeeping report with date picker
function generateHousekeepingReport() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Generate Housekeeping Report', 'Enter the date for the report in YYYY-MM-DD format\n(Example: 2025-06-24):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    const dateInput = response.getResponseText().trim();
    if (!dateInput) {
      ui.alert('Error', 'Please enter a valid date.', ui.ButtonSet.OK);
      return;
    }
    try {
      ui.alert('Processing...', `Generating housekeeping report for ${dateInput}...`, ui.ButtonSet.OK);
      processHousekeepingReport(dateInput);
      ui.alert('Report Generated Successfully', `The housekeeping report for ${dateInput} has been generated.\n\nYou can find it in the "Housekeeping Report" sheet.`, ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('Error', `Could not generate the report:\n\n${error.message}\n\nPlease verify:\n• Date format is YYYY-MM-DD\n• "Housekeeping" sheet exists with data.`, ui.ButtonSet.OK);
      debugLog('Error in housekeeping report:', error.message);
    }
  }
}

// QUICK FUNCTION: Today's housekeeping report
function todaysHousekeepingReport() {
  const today = new Date();
  const formattedDate = `${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert('Processing...', `Generating housekeeping report for today (${formattedDate})...`, ui.ButtonSet.OK);
    processHousekeepingReport(formattedDate);
    ui.alert('Report Generated Successfully', `Today's housekeeping report has been generated.\n\nYou can find it in the "Housekeeping Report" sheet.`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', `Could not generate the report:\n\n${error.message}`, ui.ButtonSet.OK);
    debugLog('Error in today\'s housekeeping report:', error.message);
  }
}

// PROCESSING FUNCTION: Handles the logic for daily housekeeping reports
function processHousekeepingReport(dateText) {
  const date = new Date(dateText + 'T00:00:00');
  if (isNaN(date.getTime())) {
    throw new Error('Invalid date format. Use YYYY-MM-DD');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const housekeepingSheet = ss.getSheetByName('Housekeeping');
  if (!housekeepingSheet) throw new Error('Could not find "Housekeeping" sheet');
  
  const tasks = getHousekeepingDataForDate(housekeepingSheet, date);
  
  let reportSheet = ss.getSheetByName('Housekeeping Report');
  if (reportSheet) reportSheet.clear();
  else reportSheet = ss.insertSheet('Housekeeping Report');
  
  generateDailyHousekeepingReportInSheet(reportSheet, tasks, date);
}

// DATA FUNCTION: Gets housekeeping tasks for a specific date
function getHousekeepingDataForDate(dataSheet, queryDate) {
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());
  const tasks = [];

  const colCabin = headers.indexOf('Cabin Names');
  const colDate = headers.indexOf('Date');
  const colTime = headers.indexOf('Time');
  const colTask = headers.indexOf('Task');
  
  if ([colCabin, colDate, colTime, colTask].includes(-1)) {
    throw new Error('Could not find all required columns in "Housekeeping" sheet: Cabin Names, Date, Time, Task.');
  }

  const queryDateISO = queryDate.toISOString().slice(0, 10);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const taskDateValue = row[colDate];
    if (!taskDateValue) continue;

    const taskDate = (taskDateValue instanceof Date) ? taskDateValue : parseAmericanDate(taskDateValue.toString());
    if (!taskDate || isNaN(taskDate.getTime())) continue;

    if (taskDate.toISOString().slice(0, 10) === queryDateISO) {
      tasks.push({
        cabinName: row[colCabin] || 'N/A',
        time: (row[colTime] instanceof Date) ? Utilities.formatDate(row[colTime], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
, 'HH:mm') : (row[colTime] || 'Any time'),
        task: row[colTask] || 'No description'
      });
    }
  }
  tasks.sort((a, b) => a.time.localeCompare(b.time));
  return tasks;
}

// SHEET GENERATION: Creates the daily housekeeping report sheet
function generateDailyHousekeepingReportInSheet(sheet, tasks, date) {
  sheet.clear();
  const formattedDate = formatReadableDate(date);
  sheet.getRange(1, 1).setValue(`DAILY HOUSEKEEPING REPORT - ${formattedDate}`);
  sheet.getRange(1, 1, 1, 3).merge().setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold').setBackground('#00acc1').setFontColor('white');
  
  if (tasks.length === 0) {
      sheet.getRange(3, 1).setValue('No housekeeping tasks scheduled for this date.');
      return;
  }
  
  const headers = ['Time', 'Cabin Name', 'Task to Perform'];
  sheet.getRange(3, 1, 1, 3).setValues([headers]);
  const detailRows = tasks.map(task => [task.time, task.cabinName, task.task]);
  sheet.getRange(4, 1, detailRows.length, 3).setValues(detailRows);
  
  formatDailyHousekeepingSheet(sheet, detailRows.length);
}

// FORMATTING: Styles the daily housekeeping report sheet
function formatDailyHousekeepingSheet(sheet, numRows) {
    const headerRange = sheet.getRange(3, 1, 1, 3);
    headerRange.setFontWeight('bold').setBackground('#00838f').setFontColor('white');
    if (numRows > 0) {
        const dataRange = sheet.getRange(3, 1, numRows + 1, 3);
        dataRange.setBorder(true, true, true, true, true, true);
        sheet.getRange(4, 1, numRows, 1).setHorizontalAlignment('center');
        for (let i = 0; i < numRows; i++) {
            if (i % 2 === 1) sheet.getRange(4 + i, 1, 1, 3).setBackground('#e0f7fa');
        }
    }
    sheet.autoResizeColumns(1, 3);
    sheet.setColumnWidth(3, 400);
}

// --- Housekeeping Outlook Functions ---

// MAIN FUNCTION: Generate housekeeping outlook report
// MAIN FUNCTION: Generate housekeeping outlook report
function generateHousekeepingOutlook() {
  const ui = SpreadsheetApp.getUi();

  // PASO 1: Preguntar por la fecha de inicio
  const dateResponse = ui.prompt('Generate Housekeeping Outlook', 'Enter the START date for the outlook period in YYYY-MM-DD format:', ui.ButtonSet.OK_CANCEL);
  if (dateResponse.getSelectedButton() !== ui.Button.OK || !dateResponse.getResponseText()) {
    return; // El usuario canceló o no ingresó nada
  }
  const startDateText = dateResponse.getResponseText().trim();

  // PASO 2: Preguntar por el número de días
  const daysResponse = ui.prompt('Generate Housekeeping Outlook', 'Enter the number of days for the outlook (e.g., 7 for a week):', ui.ButtonSet.OK_CANCEL);
  if (daysResponse.getSelectedButton() !== ui.Button.OK || !daysResponse.getResponseText()) {
    return; // El usuario canceló o no ingresó nada
  }
  const numberOfDays = parseInt(daysResponse.getResponseText().trim());

  // Validación de los datos ingresados
  if (isNaN(numberOfDays) || numberOfDays < 1 || numberOfDays > 30) {
    ui.alert('Error', 'Please enter a valid number of days (1-30).', ui.ButtonSet.OK);
    return;
  }

  try {
    const startDate = new Date(startDateText + 'T00:00:00');
    if (isNaN(startDate.getTime())) {
      throw new Error(`Invalid date format: "${startDateText}". Please use YYYY-MM-DD.`);
    }

    ui.alert('Processing...', `Generating housekeeping outlook for ${numberOfDays} days starting from ${startDateText}...`, ui.ButtonSet.OK);
    
    // Crear el rango de fechas a partir de la fecha de inicio proporcionada
    const dateRange = [];
    for (let i = 0; i < numberOfDays; i++) {
      const currentDate = new Date(startDate);
      currentDate.setDate(startDate.getDate() + i);
      dateRange.push(currentDate);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const housekeepingSheet = ss.getSheetByName('Housekeeping');
    if (!housekeepingSheet) throw new Error('Could not find "Housekeeping" sheet');
    
    const tasksByDay = getHousekeepingDataForRange(housekeepingSheet, dateRange);
    
    let reportSheet = ss.getSheetByName('Housekeeping Outlook');
    if (reportSheet) reportSheet.clear();
    else reportSheet = ss.insertSheet('Housekeeping Outlook');
    
    generateHousekeepingOutlookInSheet(reportSheet, tasksByDay, dateRange);
    
    ui.alert('Outlook Generated', `Housekeeping outlook for ${numberOfDays} days has been generated successfully.`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', `Could not generate the outlook:\n\n${error.message}`, ui.ButtonSet.OK);
    debugLog('Error in housekeeping outlook:', error.message);
  }
}

// DATA FUNCTION: Gets housekeeping tasks for a date range
function getHousekeepingDataForRange(dataSheet, dateRange) {
    const data = dataSheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const tasksByDay = new Map();

    const colCabin = headers.indexOf('Cabin Names');
    const colDate = headers.indexOf('Date');
    const colTime = headers.indexOf('Time');
    const colTask = headers.indexOf('Task');
  
    if ([colCabin, colDate, colTime, colTask].includes(-1)) {
      throw new Error('Could not find all required columns in "Housekeeping" sheet.');
    }

    const dateRangeISO = new Set(dateRange.map(d => d.toISOString().slice(0, 10)));

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const taskDateValue = row[colDate];
        if (!taskDateValue) continue;

        const taskDate = (taskDateValue instanceof Date) ? taskDateValue : parseAmericanDate(taskDateValue.toString());
        if (!taskDate || isNaN(taskDate.getTime())) continue;

        const taskDateISO = taskDate.toISOString().slice(0, 10);

        if (dateRangeISO.has(taskDateISO)) {
            if (!tasksByDay.has(taskDateISO)) {
                tasksByDay.set(taskDateISO, []);
            }
            tasksByDay.get(taskDateISO).push({
                cabinName: row[colCabin] || 'N/A',
                time: (row[colTime] instanceof Date) ? Utilities.formatDate(row[colTime], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
, 'HH:mm') : (row[colTime] || 'Any time'),
                task: row[colTask] || 'No description'
            });
        }
    }
    tasksByDay.forEach(tasks => tasks.sort((a, b) => a.time.localeCompare(b.time)));
    return tasksByDay;
}


// SHEET GENERATION: Creates the housekeeping outlook report (FINAL CORRECTED VERSION)
function generateHousekeepingOutlookInSheet(sheet, tasksByDay, dateRange) {
    sheet.clear();
    
    if (!dateRange || dateRange.length === 0) {
        sheet.getRange(1, 1).setValue('HOUSEKEEPING OUTLOOK');
        sheet.getRange(3, 1).setValue('A valid date range was not provided.');
        return;
    }

    const startDate = formatReadableDate(new Date(dateRange[0]));
    const endDate = formatReadableDate(new Date(dateRange[dateRange.length - 1]));
    
    sheet.getRange(1, 1).setValue(`HOUSEKEEPING OUTLOOK`);
    sheet.getRange(2, 1).setValue(`Period: ${startDate} to ${endDate}`);
    sheet.getRange(1, 1, 1, 3).merge().setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold').setBackground('#00acc1').setFontColor('white');
    sheet.getRange(2, 1, 1, 3).merge().setHorizontalAlignment('center').setFontSize(12).setFontWeight('bold');

    let currentRow = 4;
    
    dateRange.forEach(dateElement => {
        // ===== LA SOLUCIÓN DEFINITIVA =====
        // Se "re-hidrata" el elemento para asegurar que sea un objeto de Fecha real.
        const date = new Date(dateElement);
        // ===================================

        const dateISO = date.toISOString().slice(0, 10);
        const tasks = tasksByDay.get(dateISO) || [];
        const dayHeader = formatReadableDate(date);
        
        sheet.getRange(currentRow, 1, 1, 3).merge().setValue(dayHeader).setFontWeight('bold').setBackground('#b2ebf2').setHorizontalAlignment('center');
        currentRow++;

        if (tasks.length > 0) {
            const headers = ['Time', 'Cabin Name', 'Task'];
            sheet.getRange(currentRow, 1, 1, 3).setValues([headers]).setFontWeight('bold').setBackground('#4dd0e1');
            currentRow++;
            const rows = tasks.map(t => [t.time, t.cabinName, t.task]);
            sheet.getRange(currentRow, 1, rows.length, 3).setValues(rows);
            currentRow += rows.length;
        } else {
            sheet.getRange(currentRow, 1, 1, 3).merge().setValue('No activities assigned').setFontStyle('italic');
            currentRow++;
        }
        currentRow++;
    });

    sheet.autoResizeColumns(1, 3);
    sheet.setColumnWidth(3, 400);
}