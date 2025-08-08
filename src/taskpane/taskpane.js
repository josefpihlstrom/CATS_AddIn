/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
import { applyThickOutlineBorder, getDaysOfWeek } from "./formatting";
import { applyThinInternalBorder } from "./formatting";
import { setColumnWidths } from "./formatting";
import { writeInCell } from "./formatting";
import { getHolidays } from "./formatting";
import { writeHolidaysToExcel } from "./formatting";
import { formulaInCell } from "./formatting";
import { getColumnLetter } from "./formatting";
import { resetSheet } from "./formatting";
import { findRow } from "./formatting";
import { clearContentsOnly } from "./formatting";
import { readCell } from "./formatting";

function populateYears() {
    const yearSelect = document.getElementById("yearSelect");
    const currentYear = new Date().getFullYear();
    for (let i = 0; i <= 5; i++) {
        const year = currentYear + i;
        const option = document.createElement("option");
        option.value = year;
        option.text = year;
        yearSelect.appendChild(option);
    }
}

function setCurrentMonth() {
    const monthSelect = document.getElementById("monthSelect");
    const months = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];
    const currentMonth = new Date().getMonth();
    months.forEach((month, index) => {
        const option = document.createElement("option");
        option.value = month;
        option.text = month;
        if (index === currentMonth) {
            option.selected = true;
        }
        monthSelect.appendChild(option);
    });
}



Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runButton").addEventListener("click", handleRunClick);
        populateYears();
        setCurrentMonth();
        document.querySelector('input[value="Update"]').checked = true;
    }
});



function handleRunClick() {
    const selectedMode = document.querySelector('input[name="mode"]:checked').value;

    if (selectedMode === "New") {
        runNewScript();
    } else if (selectedMode === "Update") {
        runUpdateScript();
    }
}


function runNewScript() {
    Excel.run(async (context) => {
        console.log("Running New script...");
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const yearSelected = yearSelect.value;
        const monthSelect = document.getElementById("monthSelect");
        const monthNumber = monthSelect.selectedIndex; // 0-based index

        // Clears all data in sheet
        await resetSheet(context,sheet);

        // Builds all the borders in sheet
        applyThinInternalBorder(context,sheet,"I2:AM18");
        applyThickOutlineBorder(context,sheet,"B2:AM16");
        applyThickOutlineBorder(context,sheet,"I2:AM16");
        applyThickOutlineBorder(context,sheet,"B2:AM2");
        applyThickOutlineBorder(context,sheet,"I2:AM4");
        applyThickOutlineBorder(context,sheet,"B17:AM17");
        applyThickOutlineBorder(context,sheet,"I17:AM17");
        applyThickOutlineBorder(context,sheet,"I18:AM18");
        applyThickOutlineBorder(context,sheet,"G18:H20");
        applyThinInternalBorder(context,sheet,"G18:H20");

        // Sets all column widths in sheet
        setColumnWidths(context,sheet,"B:H",60);
        setColumnWidths(context,sheet,"F:G",100);
        setColumnWidths(context,sheet,"I:AM",25);

        // Writes all standard texts in sheet
        writeInCell(context,sheet,"B2","CATS");
        writeInCell(context,sheet,"B3","Send.CCtr");
        writeInCell(context,sheet,"C3","ActTyp");
        writeInCell(context,sheet,"D3","Rec.Order");
        writeInCell(context,sheet,"E3","Trip");
        writeInCell(context,sheet,"F3","Confirmation Text");
        writeInCell(context,sheet,"G3","Internal Time");
        writeInCell(context,sheet,"H3","Total");
        writeInCell(context,sheet,"F5","Education/Training");
        writeInCell(context,sheet,"G5","20")
        writeInCell(context,sheet,"F6","Admin/Line work");
        writeInCell(context,sheet,"G6","21");
        writeInCell(context,sheet,"F7","Meetings/Events");
        writeInCell(context,sheet,"G7","22");
        writeInCell(context,sheet,"F8","Absence");
        writeInCell(context,sheet,"G8","23");
        writeInCell(context,sheet,"B17","EndOfList")
        writeInCell(context,sheet,"H17","Remaining")
        writeInCell(context,sheet,"G18","New Flex")
        formulaInCell(context,sheet,"H18",`=SUM(I18:AM18)`)
        writeInCell(context,sheet,"G19","Saved Flex")
        writeInCell(context,sheet,"H19","0")
        writeInCell(context,sheet,"G20","Result")
        formulaInCell(context,sheet,"H20",`=SUM(H18:H19)`)

        // Loops through and writes all summary formulas in column H
        for (let rowIndex = 4; rowIndex <=16; rowIndex++) {
            const formula = `=SUM(I${rowIndex}:AM${rowIndex})`;
            await formulaInCell(context,sheet,`H${rowIndex}`,formula)
        }

        // Loops through and writes all summary formulas in row 17
        for (let colIndex = 8; colIndex <= 38; colIndex++) {
            const colLetter = getColumnLetter(colIndex);
            const formula = `=${colLetter}4-SUM(${colLetter}5:${colLetter}16)`;
            await formulaInCell(context, sheet, `${colLetter}17`, formula);
        }

        // Fetches all swedish holidays and writes them to sheet
        const holidays = await getHolidays(yearSelected);
        writeHolidaysToExcel(context,sheet,holidays,20);

        // Loops through and writes days, checking if days are weekends or holidays and updates accordingly
        getDaysOfWeek(context,sheet,yearSelected,monthNumber,holidays,"17");

        // Updates sheet
        await context.sync();

    });
}

function runUpdateScript() {
    console.log("Running Update script...");
    Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const yearSelected = yearSelect.value;
        const monthSelect = document.getElementById("monthSelect");
        const monthNumber = monthSelect.selectedIndex; // 0-based index

        // Find position of start and end to be used later
        const startRow = await findRow(context,sheet,"B1:B100", "CATS") +3 ;
        const endRow = await findRow(context, sheet, "B1:B100", "EndOfList");

        // Delete all recorded hours for the month
        await clearContentsOnly(context,sheet, `I${startRow}:AM${endRow-1}`);

        // fetch holidays
        const holidays = await getHolidays(yearSelected);

        // Update days of the month
        getDaysOfWeek(context,sheet,yearSelected,monthNumber,holidays,endRow);

        // Move value of Flex Result cell to Saved Cell
        const savedFlex = await readCell(context,sheet,`H${endRow+3}`);
        writeInCell(context,sheet,`H${endRow+2}`,savedFlex);

        // Delete all recorded flex for the month
        await clearContentsOnly(context,sheet, `I${endRow+1}:AM${endRow+1}`);

        // Updates sheet
        await context.sync();

    });
}

