/**
 * This project was developed to conduct data analysis 
 * for RCBD experimental data with replications. 
 * It performs the following: 
 * - ANOVA assumprion checks (Normality and homogeneity), 
 * - A Two-Factor ANOVA with Replications, 
 * - Test statistical significance, 
 * - Conduct result interpretation on Google Sheets.
 * 
 * @author: Wondwossen B.
 */
// Centralized color palette
const COLOR_PALETTE = {
  header: '#294189',    // Dark blue
  subHeader: '#c2ceef',
  thickBorder: '#17254f',
  configCellBg: '#ccccff', // Light Blue
  significant: '#00ff00', // Green
  ns: '#ff8080',        // Red
  warning: '#ffff80',   // Yellow
  neutral: '#95a5a6'    // Gray
};

// 95% CI - Confidence level
const CONFIDENCE_LEVEL = 0.95;

/**
 * Custom menu setup for the
 * RCBD-With-Reps ANOVA functionalities.
 * 
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  const prepSubMenu = ui.createMenu('Prepare Data')
    .addItem('Calculate All Metrics', 'prepareDataForChecks');

  const assumptionCheckMenu = ui.createMenu('ANOVA Assumptions Check')
    .addItem('Restructure Data', 'restructureData')
    .addSubMenu(prepSubMenu)
    .addItem('Generate Charts', 'createCharts');

  ui.createMenu('RCBD-With-Reps ANOVA')
    .addSubMenu(assumptionCheckMenu)
    .addItem('Run ANOVA', 'generateANOVA')
    .addToUi();
}

/**
 * Converts wide-format data (Block, T1, T2, ..., Tn) 
 * to long format (Block, Treatment, Germination).
 * Handles dynamic blocks, treatments, and replications.
 */
function restructureData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  const targetSheetName = sourceSheet.getName() + " - NH Checks";
  let targetSheet = ss.getSheetByName(targetSheetName);
  
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
  } else {
    targetSheet.clear();
  }

  // Set headers for long format
  targetSheet.getRange("A1:C1")
    .setValues([["Block", "Treatment", "Result"]])
    .setFontWeight('bold');

  // Get all data from source sheet
  const [header, ...data] = sourceSheet.getDataRange().getValues();

  // Extract treatment names (assumes first column is "Block")
  const treatments = header.slice(1);

  // Restructure data
  const output = data.flatMap(row => {
    const block = row[0];
    return treatments.map((treatment, idx) => [
      block,
      treatment,
      row[idx + 1] // +1 to skip Block column
    ]);
  });

  // Write to target sheet
  targetSheet.getRange(2, 1, output.length, 3).setValues(output);
  targetSheet.autoResizeColumns(1, 3);
}

/**
 * Function that runs all the calculations to 
 * prepare the data for ANOVA assumptions' checks.
 * 
 */
function prepareDataForChecks() {
  computeOverallMean();
  computeBlockMeans();
  computeTreatmentMeans();
  computeFittedValues();
  computeResiduals();
  computeSortedResiduals();
  computePercentiles();
  computeZScores();
  formatAssumptionCheckSheet();
}

/**
 * Function that calculates overall 
 * mean of treatments.
 * 
 */
function computeOverallMean() {
  const sheet = getCurrentNhSheet();
  const range = sheet.getRange("C2:C");
  const flattenedValues = range.getValues().flat().filter(v => v !== "");
  const overallSum = flattenedValues.reduce((a, b) => a + b, 0);
  const overallMean = overallSum / flattenedValues.length;
  
  sheet.getRange("L1")
    .setValue("Overall Treatment Mean")
    .setWrap(true)
    .setFontWeight("bold")
    .setBorder(false, false, true, false, false, false);
  sheet.getRange("L2")
    .setValue(overallMean);
}

/**
 * Function that calculates block means 
 * and fills them in the 'Block Means' 
 * column.
 * 
 */
function computeBlockMeans() {
  const sheet = getCurrentNhSheet();
  const dataRange = sheet.getRange("A2:C" + sheet.getLastRow());
  const data = dataRange.getValues();
  
  const blockMeans = data.map(row => {
    const block = row[0];
    const blockData = data.filter(r => r[0] === block);
    return blockData.reduce((sum, r) => sum + r[2], 0) / blockData.length;
  });

  sheet.getRange("D1")
    .setValue("Block Means");
  sheet.getRange(2, 4, blockMeans.length, 1)
    .setValues(blockMeans.map(m => [m]));
}

/**
 * Function that calculates treatment means 
 * and fills them in the 'Treatment Means' 
 * column.
 * 
 */
function computeTreatmentMeans() {
  const sheet = getCurrentNhSheet();
  const dataRange = sheet.getRange("A2:C" + sheet.getLastRow());
  const data = dataRange.getValues();
  
  const treatmentMeans = data.map(row => {
    const treatment = row[1];
    const treatmentData = data.filter(r => r[1] === treatment);
    return treatmentData.reduce((sum, r) => sum + r[2], 0) / treatmentData.length;
  });

  sheet.getRange("E1")
    .setValue("Treatment Means");
  sheet.getRange(2, 5, treatmentMeans.length, 1)
    .setValues(treatmentMeans.map(m => [m]));
}

/**
 * Function that calculates fitted/predicted values 
 * and fills them in the 'Fitted Values' 
 * column.
 * 
 */
function computeFittedValues() {
  const sheet = getCurrentNhSheet();
  const overallMean = sheet.getRange("L2").getValue();
  const lastRow = sheet.getLastRow();
  
  const fittedValues = sheet.getRange("D2:E" + lastRow).getValues()
    .map(([blockMean, treatmentMean]) => 
      blockMean + treatmentMean - overallMean
    );

  sheet.getRange("F1")
    .setValue("Fitted Values");
  sheet.getRange(2, 6, fittedValues.length, 1)
    .setValues(fittedValues.map(v => [v]));
}

/**
 * Function that calculates residuals by subtructing  
 * fitted values from observed values 
 * and fills them in the 'Residuals' column.
 * 
 */
function computeResiduals() {
  const sheet = getCurrentNhSheet();
  const lastRow = sheet.getLastRow();
  const results = sheet.getRange("C2:C" + lastRow).getValues().flat();
  const fittedValues = sheet.getRange("F2:F" + lastRow).getValues().flat();
  
  const residuals = results.map((g, i) => g - fittedValues[i]);

  sheet.getRange("G1")
    .setValue("Residuals");
  sheet.getRange(2, 7, residuals.length, 1)
    .setValues(residuals.map(r => [r]));
}

/**
 * Function that sorts residuals by ascending order 
 * and fills them in the 'Sorted Residuals' column.
 * 
 */
function computeSortedResiduals() {
  const sheet = getCurrentNhSheet();
  const residuals = sheet.getRange("G2:G" + sheet.getLastRow()).getValues().flat();
  const sorted = [...residuals].sort((a, b) => a - b);
  
  sheet.getRange("H1")
    .setValue("Sorted Residuals");
  sheet.getRange(2, 8, sorted.length, 1)
    .setValues(sorted.map(r => [r]));
}

/**
 * Function that calculates percentiles which will be 
 * used for Q-Q Plotting to check normality 
 * and fills them in the 'Percentiles' 
 * column.
 * 
 */
function computePercentiles() {
  const sheet = getCurrentNhSheet();
  const n = sheet.getLastRow() - 1;
  const percentiles = Array.from({length: n}, (_, i) => (i + 0.5) / n);
  
  sheet.getRange("I1")
    .setValue("Percentiles");
  sheet.getRange(2, 9, percentiles.length, 1)
    .setValues(percentiles.map(p => [p]));
}

/**
 * A custom function that calculates Z-scores which will be 
 * used for checking normality and fills them in the 'Z-Scores' 
 * column.
 * 
 */
function computeZScores() {
  const sheet = getCurrentNhSheet();
  const lastRow = sheet.getRange("I:I").getLastRow();
  
  // Get percentile values from I2 to last row
  const percentiles = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  
  // Prepare array to hold z-scores
  const zScores = percentiles.map(row => {
    const p = row[0];
    if (typeof p === 'number' && p > 0 && p < 1) {
      // Use NORM.S.INV formula in cell to get z-score
      return [sheet.getRange("J1").setFormula(`=NORM.S.INV(${p})`).getValue()];
    } else {
      return [null]; // invalid or empty percentile
    }
  });

  // Write column name
  sheet.getRange("J1").setValue("Z-Scores");

  // Write z-scores to column J starting at J2 
  sheet.getRange(2, 10, zScores.length, 1).setValues(zScores);
}

// New formatting function
function formatAssumptionCheckSheet() {
  const sheet = getCurrentNhSheet();
  sheet.getRange("A1:J1")
    .setFontWeight("bold")
    .setWrap(true)
    .setBorder(false, false, true, false, false, false);
  sheet.getDataRange().setHorizontalAlignment('center');
}

// ====================== CHART CREATION FUNCTIONS ======================
/**
 * Function that runs the functions to 
 * create Q-Q Plot and Scatter Chart which is 
 * used for testing Homogeneity.
 * 
 */
function createCharts() {
  createQQPlot();
  createResidualsVsFitted();
}

/**
 * Function to create Q-Q Plot for 
 * testing normality.
 * 
 */
function createQQPlot() {
  const sheet = getCurrentNhSheet();
  const lastRow = sheet.getLastRow();

  // Get data ranges for Q-Q plot
  const zScoresRange = sheet.getRange("J2:J" + lastRow);
  const sortedResidualsRange = sheet.getRange("H2:H" + lastRow);

  // Build and position Q-Q plot
  const qqChart = sheet.newChart()
    .asScatterChart()
    .addRange(zScoresRange)
    .addRange(sortedResidualsRange)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(0)
    .setTitle('Q-Q Plot of Residuals')
    .setXAxisTitle('Theoretical Quantiles (Z-Scores)')
    .setYAxisTitle('Sample Quantiles (Sorted Residuals)')
    .setPosition(6, 12, 0, 0) // Row 7, Column M
    .setOption('series.0.dataLabel', 'none')
    .setOption('legend.position', 'none')
    .setOption('hAxis.gridlines.count', 5)
    .setOption('vAxis.gridlines.count', 5)
    .build();

  sheet.insertChart(qqChart);
}

/**
 * Function to create Scatter Chart which is 
 * used for testing Homogeneity.
 * 
 */
function createResidualsVsFitted() {
  const sheet = getCurrentNhSheet();
  const lastRow = sheet.getLastRow();

  // Get data ranges for Residuals plot
  const fittedRange = sheet.getRange("F2:F" + lastRow);
  const residualsRange = sheet.getRange("G2:G" + lastRow);

  // Build and position Residuals plot
  const resChart = sheet.newChart()
    .asScatterChart()
    .addRange(fittedRange)
    .addRange(residualsRange)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(0)
    .setTitle('Residuals vs Fitted Values')
    .setXAxisTitle('Fitted Values')
    .setYAxisTitle('Residuals')
    .setPosition(26, 12, 0, 0) // Row 27, Column M
    .setOption('trendlines', { 0: { 
      type: 'linear',
      color: 'red',
      lineWidth: 1,
      opacity: 0.5 
    }})
    .setOption('series.0.dataLabel', 'none')
    .setOption('legend.position', 'none')
    .build();

  sheet.insertChart(resChart);
}

// ====================== TWO-FACTOR ANOVA GENERATION ======================
/**
 * Generates summary table of descriptive statistics and 
 * a Two-Factor ANOVA table in a new sheet 
 * with some formatting.
 * 
 */
function generateANOVA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();
  const rawDataSheetName = getRawDataSheetNameFromCurrent(currentSheet);
  const targetSheetName = rawDataSheetName + " - ANOVA";
  const nhSheetName = rawDataSheetName + " - NH Checks";
  
  const nhSheet = ss.getSheetByName(nhSheetName);
  if (!nhSheet) throw new Error("Run 'ANOVA Assumptions Check' first");
  
  // Create or clear ANOVA sheet
  let anovaSheet = ss.getSheetByName(targetSheetName);
  if (!anovaSheet) anovaSheet = ss.insertSheet(targetSheetName);
  else anovaSheet.clear();

  // Get raw data and design parameters
  const rawData = nhSheet.getRange("A2:C" + nhSheet.getLastRow()).getValues();
  const blocks = [...new Set(rawData.map(row => row[0]))];
  const treatments = [...new Set(rawData.map(row => row[1]))];
  const b = blocks.length, t = treatments.length, r = rawData.filter(row => 
    row[0] === blocks[0] && row[1] === treatments[0]).length;

  // Add titles
  anovaSheet.getRange("A1").setValue("Statistical Analysis for RCBD with Replication")
    .setFontWeight("bold");

  // Color formatting sheet title
  setContrastColors(anovaSheet.getRange("A1:E1"), COLOR_PALETTE.header);
  
  let currentRow = 3;
  
  // Generate descriptive statistics headers
  const headerRow = ["SUMMARY", ...treatments, "Total"];
  anovaSheet.getRange(currentRow, 1, 1, headerRow.length)
    .setValues([headerRow])
    .setFontStyle("italic")
    .setFontWeight("bold");
  // Change current row
  currentRow++;

  // Generate block summary tables
  blocks.forEach((block, blockIdx) => {
    // Add block label
    anovaSheet.getRange(currentRow, 1)
      .setValue(block)
      .setFontStyle("italic");
    currentRow++;

    const blockData = rawData.filter(row => row[0] === block);

    // Calculate statistics
    const stats = {
      count: treatments.map(t => 
        blockData.filter(row => row[1] === t).length),
      sum: treatments.map(t => 
        blockData.filter(row => row[1] === t).reduce((a, v) => a + v[2], 0)),
      avg: treatments.map(t => {
        const data = blockData.filter(row => row[1] === t).map(r => r[2]);
        return data.reduce((a, v) => a + v, 0) / data.length;
      }),
      var: treatments.map(t => {
        const data = blockData.filter(row => row[1] === t).map(r => r[2]);
        const mean = data.reduce((a, v) => a + v, 0) / data.length;
        return data.reduce((a, v) => a + Math.pow(v - mean, 2), 0) / (data.length - 1);
      })
    };

    // Add totals column
    stats.count.push(stats.count.reduce((a, v) => a + v, 0));
    stats.sum.push(stats.sum.reduce((a, v) => a + v, 0));
    stats.avg.push(stats.sum[stats.sum.length-1] / stats.count[stats.count.length-1]);
    stats.var.push(
      blockData.reduce((sum, row) => sum + Math.pow(row[2] - stats.avg[stats.avg.length-1], 2), 0) /
      (blockData.length - 1)
    );

    // Build table
    const tableData = [
      ["Count", ...stats.count],
      ["Sum", ...stats.sum],
      ["Average", ...stats.avg],
      ["Variance", ...stats.var]
    ];
    
    // Write data to table
    anovaSheet.getRange(currentRow, 1, 4, tableData[0].length)
      .setValues(tableData)
      .setBorder(true, false, false, false, false, false, COLOR_PALETTE.thickBorder, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    anovaSheet.getRange(currentRow + 1, 2, 3, tableData[0].length)
      .setNumberFormat("0.00");
      
    currentRow += 5; // 4 data rows + 1 spacer
  });

  // Generate Total table
  anovaSheet.getRange(currentRow, 1)
    .setValue("Total")
    .setFontWeight("bold")
    .setFontStyle("italic");
  currentRow++;
  
  const totalStats = {
    count: treatments.map(t => 
      rawData.filter(row => row[1] === t).length),
    sum: treatments.map(t => 
      rawData.filter(row => row[1] === t).reduce((a, v) => a + v[2], 0)),
    avg: treatments.map(t => {
      const data = rawData.filter(row => row[1] === t).map(r => r[2]);
      return data.reduce((a, v) => a + v, 0) / data.length;
    }),
    var: treatments.map(t => {
      const data = rawData.filter(row => row[1] === t).map(r => r[2]);
      const mean = data.reduce((a, v) => a + v, 0) / data.length;
      return data.reduce((a, v) => a + Math.pow(v - mean, 2), 0) / (data.length - 1);
    })
  };

  const totalTableData = [
    ["Count", ...totalStats.count],
    ["Sum", ...totalStats.sum],
    ["Average", ...totalStats.avg],
    ["Variance", ...totalStats.var]
  ];
  
  anovaSheet.getRange(currentRow, 1, 4, totalTableData[0].length)
    .setValues(totalTableData)
    .setBorder(true, false, false, false, false, false, COLOR_PALETTE.thickBorder, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  anovaSheet.getRange(currentRow + 1, 2, 3, totalTableData[0].length)
    .setNumberFormat("0.00");
  
  // Formatting first column
  anovaSheet.getRange("A:A")
    .setHorizontalAlignment("left")
    .setWrap(false);

  // Align to center except last Total header
  anovaSheet.getRange(1, 2, currentRow + 3, headerRow.length - 1)
    .setHorizontalAlignment("center");

  // Align Total header to right
  anovaSheet.getRange(1, headerRow.length, currentRow)
    .setHorizontalAlignment("right");
  
  // Generate ANOVA table after spacing
  currentRow += 8;
  generateANOVATable(anovaSheet, currentRow, rawData, blocks, treatments, b, t, r);

  // Color formatting P-value cells
  formatANOVATable(anovaSheet, currentRow);

  // Generate Statistical Interpretation
  generateStatisticalInterpretation(anovaSheet, currentRow);
}

// ====================== HELPER FUNCTION FOR ANOVA TABLE ======================
/**
 * Generates the Two-Factor ANOVA table.
 * 
 * @param {Sheet} sheet The Google Sheets Sheet object where the ANOVA table is generated.
 * @param {Number} startRow Row number where the ANOVA table starts.
 * @param {Object} rawData Observations raw data.
 * @param {Object} blocks Array of blocks data.
 * @param {Object} treatments Array of treatments data.
 * @param {Number} b Number of blocks.
 * @param {Number} t Number of treatments.
 * @param {Number} r Number of treatment replications per block.
 * @customFunction
 */
function generateANOVATable(sheet, startRow, rawData, blocks, treatments, b, t, r) {
  const { ssBlocks, ssTreatments, ssInteraction, ssError, ssTotal } = 
    calculateSSFromRaw(rawData, blocks, treatments, b, t, r);
  
  // Degrees of freedom
  const dfBlocks = b - 1;
  const dfTreatments = t - 1;
  const dfInteraction = (b - 1) * (t - 1);
  const dfError = b * t * (r - 1);
  const dfTotal = b * t * r - 1;

  // Mean squares
  const msBlocks = ssBlocks / dfBlocks;
  const msTreatments = ssTreatments / dfTreatments;
  const msInteraction = ssInteraction / dfInteraction;
  const msError = ssError / dfError;

  // F-values
  const fBlocks = msBlocks / msError;
  const fTreatments = msTreatments / msError;
  const fInteraction = msInteraction / msError;

  // Effect sizes
  const etaSqBlocks = ssBlocks / ssTotal;
  const etaSqTreatments = ssTreatments / ssTotal;
  const etaSqInteraction = ssInteraction / ssTotal;
  
  const omegaSqBlocks = (ssBlocks - dfBlocks * msError) / (ssTotal + msError);
  const omegaSqTreatments = (ssTreatments - dfTreatments * msError) / (ssTotal + msError);
  const omegaSqInteraction = (ssInteraction - dfInteraction * msError) / (ssTotal + msError);

  // Build ANOVA table with CI columns
  const anovaData = [
    ['Source', 'SS', 'df', 'MS', 'F', 'η²', 'η² 95% CI', 'ω²', 'ω² 95% CI', 'P-value', 'F crit'],
    ['Blocks', ssBlocks, dfBlocks, msBlocks, fBlocks, etaSqBlocks, '', omegaSqBlocks, '', '', ''],
    ['Treatments', ssTreatments, dfTreatments, msTreatments, fTreatments, etaSqTreatments, '', omegaSqTreatments, '', '', ''],
    ['Interaction', ssInteraction, dfInteraction, msInteraction, fInteraction, etaSqInteraction, '', omegaSqInteraction, '', '', ''],
    ['Error', ssError, dfError, msError, '', '', '', '', '', '', ''],
    ['Total', ssTotal, dfTotal, '', '', '', '', '', '', '', '']
  ];
  
  // Add ANOVA title
  sheet.getRange(startRow, 1).setValue("ANOVA: Two-Factor With Replication")
    .setFontWeight("bold")
    .setHorizontalAlignment('left');

  // Set configurable α
  const alphaCell = sheet.getRange(startRow, 9);
  sheet.getRange(startRow, 8).setValue("α:")
    .setHorizontalAlignment("right");
  alphaCell
    .setValue(0.05)
    .setNumberFormat("0.00");

  // Color format alpha value cell
  setContrastColors(alphaCell, COLOR_PALETTE.configCellBg);
  startRow++;

  // Write to sheet
  const tableRange = sheet.getRange(startRow, 1, anovaData.length, anovaData[0].length);
  const anovaLabelRange = sheet.getRange(startRow, 1, 1, anovaData[0].length);
  tableRange.setValues(anovaData);

  // Formatting Borders
  anovaLabelRange
    .setFontStyle("italic")
    .setBorder(false, false, true, false, null, null, COLOR_PALETTE.thickBorder, SpreadsheetApp.BorderStyle.SOLID);
  tableRange.setBorder(true, true, true, true, null, null, COLOR_PALETTE.thickBorder, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Number formatting
  sheet.getRange(startRow, 2, anovaData.length - 1, 1).setNumberFormat("0.000"); // SS
  sheet.getRange(startRow, 4, anovaData.length - 1, 2).setNumberFormat("0.000"); // MS & F
  // η² & η² 95% CI, ω² & ω² 95% CI as percentages
  sheet.getRange(startRow, 6, anovaData.length - 1, 4).setNumberFormat("0.00%");
  sheet.getRange(startRow, 10, anovaData.length - 1, 2).setNumberFormat("0.0000"); // P-value & F crit

  // Center horizontally starting column 2
  sheet.getRange(startRow, 2, anovaData.length, anovaData[0].length).setHorizontalAlignment('center');

  // Add formulas for P-value and F crit
  [1, 2, 3].forEach((rowOffset, idx) => { // Blocks, Treatments, Interaction rows
    const row = startRow + rowOffset;
    const F = sheet.getRange(row, 5).getValue();
    
    if (isNaN(F)) {
      throw new Error(`Invalid F-value at row ${row}`);
    }

    const df1 = sheet.getRange(row, 3).getValue();
    const df2 = sheet.getRange(startRow + 4, 3).getValue(); // Error df

    try {
      // Calculate CIs
      const ci = calculateEffectSizeCI(sheet, F, df1, df2);

      // η² CI
      sheet.getRange(row, 7)
        .setValue(`[${ci.etaCI[0].toFixed(3)}, ${ci.etaCI[1].toFixed(3)}]`)
        .setNote(`95% CI for η²: ${(ci.etaCI[0]*100).toFixed(1)}% to ${(ci.etaCI[1]*100).toFixed(1)}%`);
      
      // ω² CI
      sheet.getRange(row, 9)
        .setValue(`[${ci.omegaCI[0].toFixed(3)}, ${ci.omegaCI[1].toFixed(3)}]`)
        .setNote(`95% CI for ω²: ${(ci.omegaCI[0]*100).toFixed(1)}% to ${(ci.omegaCI[1]*100).toFixed(1)}%`);

      sheet.getRange(row, 10).setFormula(`=F.DIST.RT(E${row}, C${row}, C${startRow + 4})`); // P-value
      sheet.getRange(row, 11).setFormula(`=F.INV.RT(${alphaCell.getValue()}, C${row}, C${startRow + 4})`); // F crit
    } catch (e) {
      console.error(`CI calculation failed: ${e.message}`);
      sheet.getRange(row, 7).setValue("CI Error");
      sheet.getRange(row, 9).setValue("CI Error");
    }

  });

  // Auto resizing column 2-11 at this point
  sheet.autoResizeColumns(2, 11);
}

/**
 * Retruns Sum of Squares calculations for 
 * a Two-Factor ANOVA table generation.
 * 
 * @param {Object} rawData Observations raw data.
 * @param {Object} blocks Array of blocks data.
 * @param {Object} treatments Array of treatments data.
 * @param {Number} b Number of blocks.
 * @param {Number} t Number of treatments.
 * @param {Number} r Number of treatment replications per block.
 * @returns {Object}  Sum of Squares.
 */
function calculateSSFromRaw(rawData, blocks, treatments, b, t, r) {
  let ssBlocks = 0, ssTreatments = 0, ssInteraction = 0, ssError = 0, ssTotal = 0;
  
  // 1. Calculate overall mean
  const overallMean = rawData.reduce((sum, row) => sum + row[2], 0) / rawData.length;

  // 2. Calculate block and treatment means
  const blockMeans = new Map();
  blocks.forEach(block => {
    const blockData = rawData.filter(row => row[0] === block);
    blockMeans.set(block, blockData.reduce((sum, row) => sum + row[2], 0) / blockData.length);
  });

  const treatmentMeans = new Map();
  treatments.forEach(treatment => {
    const treatmentData = rawData.filter(row => row[1] === treatment);
    treatmentMeans.set(treatment, treatmentData.reduce((sum, row) => sum + row[2], 0) / treatmentData.length);
  });

  // 3. Calculate group means and SS components
  const groupMeans = new Map();
  blocks.forEach(block => {
    treatments.forEach(treatment => {
      const groupData = rawData.filter(row => 
        row[0] === block && row[1] === treatment
      ).map(row => row[2]);
      
      const mean = groupData.reduce((a, b) => a + b, 0) / groupData.length;
      groupMeans.set(`${block}|${treatment}`, mean);
      
      // SS Error
      ssError += groupData.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0);
    });
  });

  // SS Blocks
  blocks.forEach(block => {
    ssBlocks += t * r * Math.pow(blockMeans.get(block) - overallMean, 2);
  });

  // SS Treatments
  treatments.forEach(treatment => {
    ssTreatments += b * r * Math.pow(treatmentMeans.get(treatment) - overallMean, 2);
  });

  // SS Interaction
  blocks.forEach(block => {
    treatments.forEach(treatment => {
      const interaction = groupMeans.get(`${block}|${treatment}`) 
        - blockMeans.get(block) 
        - treatmentMeans.get(treatment) 
        + overallMean;
      ssInteraction += r * Math.pow(interaction, 2);
    });
  });

  // SS Total
  ssTotal = rawData.reduce((sum, row) => sum + Math.pow(row[2] - overallMean, 2), 0);

  return { ssBlocks, ssTreatments, ssInteraction, ssError, ssTotal };
}

/**
 * Formats ANOVA table.
 * 
 * @param {Sheet} sheet The Google Sheets Sheet object containing the ANOVA table.
 * @param {Number} startRow Row number where the ANOVA table starts.
 */
function formatANOVATable(sheet, startRow) {
  // get alpha value
  const alphaVal = sheet.getRange(startRow, 9).getValue();

  // set next bigger alpha value
  const alphaNextVal = alphaVal == 0.01 ? 0.05
    : alphaVal == 0.05 ? 0.1
    : alphaVal*2;

  // color formatting
  const headerRange = sheet.getRange(`A${startRow}:E${startRow}`);
  setContrastColors(headerRange, COLOR_PALETTE.header);
  startRow++;

  [1, 2, 3].forEach((rowOffset, idx) => { // Blocks, Treatments, Interaction rows
    const row = startRow + rowOffset;
    const pValueCell = sheet.getRange(row, 10);
    const pValue = pValueCell.getValue();
    const bgColor = pValue < alphaVal ? COLOR_PALETTE.significant 
      : (pValue >= alphaVal && pValue < alphaNextVal) ? COLOR_PALETTE.warning
      : COLOR_PALETTE.ns;

    // Set the contrasted colors
    setContrastColors(pValueCell, bgColor);
  });
}

// =========================== ANOVA RESULT INTERPRETATION ========================
/**
 * Generates the statistical interpretation after ANOVA.
 * 
 * @param {Sheet} sheet           The Google Sheets Sheet object 
 *                                containing the ANOVA table 
 *                                where the statistical interpretation 
 *                                will be created.
 * @param {Number} anovaStartRow  Row number where the statistical 
 *                                interpretation table starts.
 */
function generateStatisticalInterpretation(sheet, anovaStartRow) {
  const pvalStartRow = anovaStartRow + 2;
  const statSigStartRow = anovaStartRow + 9;
  const interpretationStartRow = anovaStartRow + 14;

  // Testing Statistical Significance
  testStatisticalSignificances(sheet, pvalStartRow, statSigStartRow);

  // Interpreting the Results
  generateResultInterpretation(sheet, statSigStartRow, interpretationStartRow);
}

function testStatisticalSignificances(sheet, pvalStartRow, startRow) {
  // Add title and labels
  const statSigLabels = [
    ['Statistical Significance Tests', '', ''],
    ['Block', 'Treatment', 'Error']
  ];

  // get alpha value cell
  const anovaStart = sheet.createTextFinder("ANOVA: Two-Factor With Replication").findNext().getRow();
  const alphaCell = sheet.getRange(anovaStart, 9);
  const alphaVal = alphaCell.getValue();
  const alphaNextVal = alphaVal == 0.01 ? 0.05
    : alphaVal == 0.05 ? 0.1
    : alphaVal*2;

  // Write labels
  sheet.getRange(startRow, 1, statSigLabels.length, 3).setValues(statSigLabels);

  // Add formulas to test statistical significances
  const statSigLabelsRow = startRow + 1;
  const statSigTestRow = startRow + 2;
  const trtPvalRow = pvalStartRow + 1;
  const errPvalRow = pvalStartRow + 2;

  // get P-value cells
  const b_pcell = sheet.getRange(`J${pvalStartRow}`);
  const t_pcell = sheet.getRange(`J${trtPvalRow}`);
  const err_pcell = sheet.getRange(`J${errPvalRow}`);

  // Add to testStatisticalSignificances()
  if (!b_pcell.getValue() || !t_pcell.getValue()) {
    throw new Error("ANOVA results missing - run analysis first");
  }

  // Statistical significances for:
  // Block
  sheet.getRange(`A${statSigTestRow}`).setFormula(
    `=IF(${b_pcell.getA1Notation()}<${alphaVal},"***",IF(AND(${b_pcell.getA1Notation()}>=${alphaVal}, ${b_pcell.getA1Notation()}<${alphaNextVal}),"**","ns"))`
  );

  // Treatment
  sheet.getRange(`B${statSigTestRow}`).setFormula(
    `=IF(${t_pcell.getA1Notation()}<${alphaVal},"***",IF(AND(${t_pcell.getA1Notation()}>=${alphaVal}, ${t_pcell.getA1Notation()}<${alphaNextVal}),"**","ns"))`
  );

  // Error (Within)
  sheet.getRange(`C${statSigTestRow}`).setFormula(
    `=IF(${err_pcell.getA1Notation()}<${alphaVal},"***",IF(AND(${err_pcell.getA1Notation()}>=${alphaVal}, ${err_pcell.getA1Notation()}<${alphaNextVal}),"**","ns"))`
  );

  // Formatting and some note
  sheet.getRange(`A${startRow}:C${startRow}`)
    .setFontWeight("bold")
    .setWrap(false);

  // contrasted colors
  setContrastColors(
    sheet.getRange(`A${startRow}:C${startRow}`),
    COLOR_PALETTE.header
  );

  sheet.getRange(`C${startRow}`)
    .setNote(`Cutoff: ${alphaVal}\nCalculation: =${b_pcell.getFormula()}`);
  sheet.getRange(`A${statSigLabelsRow}:C${statSigTestRow}`)
    .setFontStyle("italic");
}

function generateResultInterpretation(sheet, statSigStartRow, startRow) {
  // Add 'Result Interpretation' title
  sheet.getRange(startRow, 1).setValue("Result Interpretation");
  
  const blkEffectStartRow = startRow + 1;
  const blkEffectRow = startRow + 2;
  const trtEffectStartRow = startRow + 4;
  const trtEffectRow = startRow + 5;
  const effSizeKeyStartRow = startRow + 8;

  // Row of stat. sig. test row and cells
  const statSigTestRow = statSigStartRow + 2;
  const bStatSigCell = sheet.getRange(`A${statSigTestRow}`);
  const tStatSigCell = sheet.getRange(`B${statSigTestRow}`);
  const errStatSigCell = sheet.getRange(`C${statSigTestRow}`);
  
  // Add 'Block Effect' label
  sheet.getRange(blkEffectStartRow, 1).setValue("Block Effect");

  // Add 'Treatment Effect' label
  sheet.getRange(trtEffectStartRow, 1).setValue("Treatment Effect");

  // Get η² values from ANOVA table
  const anovaStart = sheet.createTextFinder("ANOVA: Two-Factor With Replication").findNext().getRow();
  const blocksEta = sheet.getRange(anovaStart + 2, 6); // η² for Blocks
  const treatmentsEta = sheet.getRange(anovaStart + 3, 6); // η² for Treatments
  const interactionEta = sheet.getRange(anovaStart + 4, 6); // η² for Interaction

  // Get CI ranges
  const blocksEtaCI = sheet.getRange(anovaStart + 2, 7).getValue();
  const treatmentsEtaCI = sheet.getRange(anovaStart + 3, 7).getValue();

  // Block Effect formula
  const blockFormula = 
    `=IF(${bStatSigCell.getA1Notation()}="***",` +
    `"Statistically significant differences (η²=" & TEXT(${blocksEta.getA1Notation()},"0.00%") &` +
    `IF(${blocksEta.getA1Notation()}>=0.14,", large effect)",` +
    `IF(${blocksEta.getA1Notation()}>=0.06,", medium effect)",", small effect)")) &` +
    `" Suggests " & IF(${blocksEta.getA1Notation()}>=0.06, "important ", "negligible ") &` +
    `"blocking factor control.",` +
    `"No significant block differences")`;

  // Enhanced interpretation with error handling
  try {
    sheet.getRange(`A${blkEffectRow}`).setFormula(blockFormula);
  } catch (e) {
    console.error(`Block effect formula error: ${e.message}`);
    sheet.getRange(`A${blkEffectRow}`).setValue("Interpretation Error");
  }

  // Treatment Effect formula  
  const treatmentFormula =
    `=IF(${tStatSigCell.getA1Notation()}="***",` +
    `"Statistically significant differences (η²=" & TEXT(${treatmentsEta.getA1Notation()},"0.00%") &` +
    `IF(${treatmentsEta.getA1Notation()}>=0.14,", large effect)",` +
    `IF(${treatmentsEta.getA1Notation()}>=0.06,", medium effect)",", small effect)")) &` +
    `" Practical significance: " & IF(${treatmentsEta.getA1Notation()}>=0.06, "meaningful ", "marginal ") &` +
    `"differences.",` +
    `"No significant treatment differences")`;

  try {
    sheet.getRange(`A${trtEffectRow}`).setFormula(treatmentFormula);
  } catch (e) {
    console.error(`Treatment effect formula error: ${e.message}`);
    sheet.getRange(`A${trtEffectRow}`).setValue("Interpretation Error");
  }
  
  // Add effect size key
  sheet.getRange(effSizeKeyStartRow, 1).setValue("Effect Size Interpretation Key:")
    .setFontWeight("bold");
  const effectSizeKey = [
    ["η² ≥ 14%", "Large effect", "Practically important differences"],
    ["6% ≤ η² < 14%", "Medium effect", "Meaningful but moderate differences"],
    ["1% ≤ η² < 6%", "Small effect", "Marginal practical significance"],
    ["η² < 1%", "Negligible effect", "Likely unimportant differences"]
  ];
  sheet.getRange(effSizeKeyStartRow + 1, 1, effectSizeKey.length, effectSizeKey[0].length)
    .setValues(effectSizeKey)
    .setBorder(true, true, true, true, null, null, COLOR_PALETTE.neutral, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Add CI interpretation guide
  const ciInterpretation = [
    ["CI Range", "Interpretation"],
    ["Excludes 0", "Statistically significant effect"],
    ["Includes 0", "Effect not statistically significant"],
    ["Entirely >6%", "Practically important effect"],
    ["Partially <6%", "Uncertain practical significance"]
  ];
  
  sheet.getRange(startRow + 15, 1, ciInterpretation.length, ciInterpretation[0].length)
    .setValues(ciInterpretation)
    .setBorder(true, true, true, true, null, null, COLOR_PALETTE.neutral, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Formatting header and effect labels
  sheet.getRange(`A${startRow}:C${startRow}`)
    .setFontWeight("bold")
    .setHorizontalAlignment('left')
    .setWrap(false);

  sheet.getRange(`A${blkEffectStartRow}:B${blkEffectStartRow}`)
    .setFontStyle("italic")
    .setHorizontalAlignment('left');

  sheet.getRange(`A${trtEffectStartRow}:B${trtEffectStartRow}`)
    .setFontStyle("italic")
    .setHorizontalAlignment('left');

  // Header contrasted color formatting
  setContrastColors(
    sheet.getRange(`A${startRow}:C${startRow}`),
    COLOR_PALETTE.header
  );

  // Block effect contrasted color formatting
  setContrastColors(
    sheet.getRange(`A${blkEffectStartRow}:B${blkEffectStartRow}`),
    COLOR_PALETTE.subHeader
  );

  // Treatment effect contrasted color formatting
  setContrastColors(
    sheet.getRange(`A${trtEffectStartRow}:B${trtEffectStartRow}`),
    COLOR_PALETTE.subHeader
  );

  // Add  notes
  addResultInterpretationNotes(sheet, `B${blkEffectStartRow}`, `B${trtEffectStartRow}`);

  // Effect size key header contrasted color
  setContrastColors(
    sheet.getRange(`A${effSizeKeyStartRow}:C${effSizeKeyStartRow}`),
    COLOR_PALETTE.header
  );

  // Style first column italic
  sheet.getRange(effSizeKeyStartRow + 1, 1, effectSizeKey.length)
    .setFontStyle("italic");
}

function addResultInterpretationNotes(sheet, blkEffectLabelA1, trtEffectLabelA1) {
  sheet.getRange(blkEffectLabelA1)
    .setNote("A significant block effect suggests that the blocks are not homogeneous. It shows the experimental design has successfully controlled for some source of variation.");
  sheet.getRange(trtEffectLabelA1)
    .setNote("A significant F-statistic indicates that differences between the treatment means are statistically significant.");
}

// ====================== EFFECT SIZE CI CALCULATIONS ======================

function calculateEffectSizeCI(sheet, F, df1, df2) {
  const tempCell = sheet.getRange("ZZ1"); // Use an unlikely-to-be-used cell

  // Calculate lower F critical value
  tempCell.setFormula(`=F.INV(${(1 - CONFIDENCE_LEVEL)/2}, ${df1}, ${df2})`);
  const lowerF = tempCell.getValue();
  
  // Calculate upper F critical value
  tempCell.setFormula(`=F.INV(${1 - (1 - CONFIDENCE_LEVEL)/2}, ${df1}, ${df2})`);
  const upperF = tempCell.getValue();
  
  tempCell.clear();

  // η² CI using Smithson (2003) approximation
  const lowerEta = Math.max(0, (df1 * lowerF - df1) / (df1 * lowerF + df2));
  const upperEta = Math.min(1, (df1 * upperF - df1) / (df1 * upperF + df2));

  // ω² CI using Fidler & Thompson (2001) method
  const lowerOmega = Math.max(0, 
    (df1 * (lowerF - 1)) / (df1 * lowerF + df2 + 1));
  const upperOmega = Math.min(1, 
    (df1 * (upperF - 1)) / (df1 * upperF + df2 + 1));

  return {
    etaCI: [lowerEta, upperEta],
    omegaCI: [lowerOmega, upperOmega]
  };
}

// ====================== COLOR CONTRAST UTILITIES ======================
function setContrastColors(range, hexColor) {
  const fontColor = getContrastFontColor(hexColor);
  range
    .setBackground(hexColor)
    .setFontColor(fontColor);
}

function getContrastFontColor(hexColor) {
  // Convert hex to RGB
  let hex = hexColor.replace('#', '');
  if (hex.length === 3) {
    hex = hex.split('').map(c => c + c).join('');
  }
  
  const r = parseInt(hex.substr(0, 2), 16) / 255;
  const g = parseInt(hex.substr(2, 2), 16) / 255;
  const b = parseInt(hex.substr(4, 2), 16) / 255;

  // Calculate relative luminance (WCAG formula)
  const luminance = 0.2126 * (r <= 0.03928 ? r / 12.92 : Math.pow((r + 0.055) / 1.055, 2.4)) +
                    0.7152 * (g <= 0.03928 ? g / 12.92 : Math.pow((g + 0.055) / 1.055, 2.4)) +
                    0.0722 * (b <= 0.03928 ? b / 12.92 : Math.pow((b + 0.055) / 1.055, 2.4));

  // Choose font color based on contrast
  return luminance > 0.179 ? '#2c3e50' : '#ecf0f1';
}

// ====================== HELPER FUNCTIONS ======================
/**
 * A helper function to get the sheet where 
 * Normality and Homogeiniety is checked, 
 * contains restructured data.
 * 
 * @return Sheet with `${rawDataSheetName} - NH Checks` sheet name.
 */
function getCurrentNhSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const rawDataSheetName = getRawDataSheetNameFromCurrent(sheet);
  return ss.getSheetByName(rawDataSheetName + " - NH Checks");
}

/**
 * Returns raw data sheet name that contains 
 * treatment data before analysis.
 * 
 * @param {Sheet} currentSheet    The Google Sheets Sheet object where 
 *                                user is currently at.
 * @param {String}                Name of sheet with raw data, used for analysis.
 */
function getRawDataSheetNameFromCurrent(currentSheet) {
  const currentSheetName = currentSheet.getName();

  // check which sheet and return raw data sheet name
  if (currentSheetName.indexOf(" - ANOVA") != -1) {
    return currentSheetName.slice(0, currentSheetName.indexOf(" - ANOVA"));
  } else if (currentSheetName.indexOf(" - NH Checks") != -1) {
    return currentSheetName.slice(0, currentSheetName.indexOf(" - NH Checks"));
  } else {
    return currentSheetName;
  }
}