
/**
 * Applies a thick outline border around the range B2:AM16 in the active worksheet.
 */

export async function applyThickOutlineBorder(context, sheet, rangeAddress) {
    try {
 //       await Excel.run(async (context) => {
//            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(rangeAddress);

            const borders = range.format.borders;

            const borderItems = [
                "EdgeTop",
                "EdgeBottom",
                "EdgeLeft",
                "EdgeRight"
            ];

            for (let i = 0; i < borderItems.length; i++){
                const border = borders.getItem(borderItems[i]);
                border.style = Excel.BorderLineStyle.continuous;
                border.color = "black";
                border.weight = "Thick";
            }

            //await context.sync();
//        });
    } catch (error) {
        console.error("Error in applyThickOutlineBorder:", error.message);
    }
}

export async function applyThinInternalBorder(context, sheet,rangeAddress) {
    try {
 //       await Excel.run(async (context) => {
 //           const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(rangeAddress);

            const borders = range.format.borders;

            const borderItems = [
                "InsideHorizontal",
                "InsideVertical"
            ];

            for (let i = 0; i < borderItems.length; i++){
                const border = borders.getItem(borderItems[i]);
                border.style = Excel.BorderLineStyle.continuous;
                border.color = "gray";
                border.weight = "Thin";
            }

            //await context.sync();
 //       });
    } catch (error) {
        console.error("Error in applyThinInternalBorder:", error.message);
    }
}

export async function setColumnWidths(context, sheet, rangeAddress,width) {
    try {
 //       await Excel.run(async (context) => {
//            const sheet = context.workbook.worksheets.getActiveWorksheet();

            // Get the entire columns I through AM
            const range = sheet.getRange(rangeAddress).getEntireColumn();
            range.format.columnWidth = width;

            //await context.sync();
 //       });
    } catch (error) {
        console.error("Error in setColumnWidths:", error.message);
    }
}

export async function readCell(context, sheet, cellAddress) {
    const range = sheet.getRange(cellAddress);
    range.load("values");
    await context.sync();

    return range.values[0][0]; // Assuming it's a single cell
}

export async function writeInCell(context, sheet, rangeAddress,string) {
    const range = sheet.getRange(rangeAddress);
    range.values = [[string]];
}

export async function formulaInCell(context, sheet, rangeAddress,string) {
    const range = sheet.getRange(rangeAddress);
    range.formulas = [[string]];
}


export function getDaysOfWeek(context, sheet, year, month, holidays,endOfList) {
    const daysInMonth = new Date(year, month + 1, 0).getDate(); // month is 0-based
    const holidayDates = holidays.map(h => h.date); // assuming 'YYYY-MM-DD'

    // Fill actual days
    for (let day = 1; day <= daysInMonth; day++) {
        const dateObj = new Date(year, month, day);
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayOfWeek = dateObj.toLocaleDateString('en-US', { weekday: 'short' });

        const colIndex = 8 + (day - 1); // I = 8
        const colLetter = getColumnLetter(colIndex);

        const isWeekend = ["Sat", "Sun"].includes(dayOfWeek);
        const isHoliday = holidayDates.includes(dateStr);

        writeInCell(context, sheet, `${colLetter}2`, dayOfWeek);
        writeInCell(context, sheet, `${colLetter}3`, day);
        writeInCell(context, sheet, `${colLetter}4`, (isWeekend || isHoliday) ? 0 : 8);

        const range = sheet.getRange(`${colLetter}2:${colLetter}${endOfList+1}`);
        if (isWeekend) {
            range.format.fill.color = "#D3D3D3"; // gray
        } else if (isHoliday) {
            range.format.fill.color = "#FF9999"; // red
        } else {
            range.format.fill.color = "white"; // default
        }
    }

    // Clear remaining columns up to day 31
    for (let day = daysInMonth + 1; day <= 31; day++) {
        const colIndex = 8 + (day - 1);
        const colLetter = getColumnLetter(colIndex);

        writeInCell(context, sheet, `${colLetter}2`, "");
        writeInCell(context, sheet, `${colLetter}3`, "");
        writeInCell(context, sheet, `${colLetter}4`, "");

        const range = sheet.getRange(`${colLetter}2:${colLetter}4`);
        range.format.fill.color = "white"; // reset color
    }
}


export async function getHolidays(year) {
    try {
        const holidaysResponse = await fetch(`https://api.dagsmart.se/holidays?year=${year}&weekends=true`);
        if (!holidaysResponse.ok) throw new Error("Holiday API request failed");
        const holidays = await holidaysResponse.json();

        const bridgeDaysResponse = await fetch(`https://api.dagsmart.se/bridge-days?year=${year}&weekends=false`);
        if (!bridgeDaysResponse.ok) throw new Error("Bridge-day API request failed");
        const bridgeDaysRaw = await bridgeDaysResponse.json();

        // Tag bridge days
        const bridgeDays = bridgeDaysRaw.map(day => ({
            ...day,
            name: {
                ...day.name,
                en: "Bridge Day",
                sv: "Klämdag"
            }
        }));

        // Combine and sort
        const combined = [...holidays, ...bridgeDays].sort((a, b) => {
            return new Date(a.date) - new Date(b.date);
        });

        return combined;
    } catch (error) {
        console.error("Error fetching holidays and bridge-days:", error);
        return [];
    }
}

export function writeHolidaysToExcel(context, sheet, holidays,rowNumber) {
    let i = 0;
    sheet.getRange(`B${rowNumber}`).values = "Holidays"
    holidays.forEach((holiday) => {
        const nameEn = holiday.name?.en;
        if (nameEn !== "Saturday" && nameEn !== "Sunday") {
            const row = rowNumber + 1 + i;
            sheet.getRange(`B${row}`).values = [[holiday.date]];
            sheet.getRange(`C${row}`).values = [[holiday.name?.sv || ""]];
            i++;
        }
    });
}

export function getColumnLetter(colIndex) {
  let letter = '';
  while (colIndex >= 0) {
    letter = String.fromCharCode((colIndex % 26) + 65) + letter;
    colIndex = Math.floor(colIndex / 26) - 1;
  }
  return letter;
}

export async function resetSheet(context, sheet) {
    const range = sheet.getUsedRange();
    range.load("address"); // Load something to ensure it's accessible
    await context.sync();  // Sync to populate the range

    range.clear(Excel.ClearApplyTo.all); // Clear everything: values, formulas, formats, borders
}

export async function findRow(context, sheet, rangeAddress, string) {
    const range = sheet.getRange(rangeAddress);
    range.load("values");
    await context.sync();

    for (let i = 0; i < 100; i++) {
        const cellValue = range.values[i][0];
        if (cellValue === string) {
            console.log(`Found String at row ${i + 1}`);
            return i + 1; // Excel rows are 1-based
        }
    }

    console.error("String not found in column B rows 1–100.");
    return null;
}

export async function clearContentsOnly(context, sheet, rangeAddress) {
    const range = sheet.getRange(rangeAddress);
    range.clear(Excel.ClearApplyTo.contents); // Clears values and formulas only
}
