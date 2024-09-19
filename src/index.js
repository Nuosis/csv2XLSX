import * as XLSX from 'xlsx';

function extractDocketNumbers(text) {
  // Regular expression to match the docket number format ##-NN-######
  const docketNumberRegex = /\b\d{2}-[A-Z]{2}-\d{6}\b/g;
  
  // Use the match method to find all matches of the regex in the text
  const docketNumbers = text.match(docketNumberRegex);
  
  // Return the unique docket numbers if any were found, or an empty array
  return docketNumbers ? [...new Set(docketNumbers)] : [];
}

function splitLines(input, docketNumber) {
  // Find the index of the docket number in the string
  const docketIndex = input.indexOf(docketNumber);
  
  if (docketIndex === -1) {
    console.error('Docket number not found.');
    return;
  }

  // Start searching for commas backwards from the docket index
  let commaCount = 0;
  let splitIndex = docketIndex;
  while (commaCount < 4 && splitIndex > 0) {
    splitIndex--;
    if (input[splitIndex] === ',') {
      commaCount++;
    }
  }

  // Check if 4 commas have been found, if not, log an error
  if (commaCount < 4) {
    console.error('Not enough commas found before the docket number.');
    return;
  }

  // Find the uppercase letter immediately following the last comma
  const match = /[A-Z]/.exec(input.substring(splitIndex));
  if (match) {
    // Adjust split index to the position of the uppercase letter
    splitIndex -= match.index;
  } else {
    console.error('Uppercase letter not found after the fourth comma.');
    return;
  }

  // Split the input at the calculated index
  const line1 = input.substring(0, splitIndex).trim();
  const line2 = input.substring(splitIndex).trim();

  return [line1, line2];
}

function checkDocketNumbersInCleanedData(docketNumbers, cleanedData) {
  const misplaced = [];
  const missing = [];
  let repairedData = cleanedData;

  // Split cleanedData into lines
  const cleanedDataLines = cleanedData.split(/\r\n|\n|\r/);

  // Iterate through each docket number
  docketNumbers.forEach(docketNumber => {
    const lineContainingDocket = cleanedDataLines.find(line => {
      const columns = line.split(',');
      // Check if docket number is in the 4th or 5th column
      const exists = columns[4] && columns[4].includes(docketNumber) ||  columns[3] && columns[3].includes(docketNumber)
      return exists;
    });

    if (lineContainingDocket) {
      // console.log("found: ",docketNumber)
      // Correctly placed, do nothing
    } else {
      console.log("error: ",docketNumber)
      const lineContainingMisplaceDocket = cleanedDataLines.find(line => {
        return line.includes(docketNumber);
      })
      console.log({lineContainingMisplaceDocket})
      // If not found in the 5th column, check if it's misplaced or completely missing
      const foundInAnyOtherLine = cleanedDataLines.some(line => line.includes(docketNumber));
      if (foundInAnyOtherLine) {
        // Split the misplaced line
        let repairedLines = splitLines(lineContainingMisplaceDocket, docketNumber);
        console.log(repairedLines)
        
        // Replace the misplaced line in repairedData with the two new lines
        if (repairedLines && repairedLines.length === 2) {
          const line1 = repairedLines[0];
          const line2 = repairedLines[1];

          // Replacing line1line2 with line1\nline2
          repairedData = repairedData.replace(lineContainingMisplaceDocket, `${line1}\n${line2}`);
        } else {
          misplaced.push(docketNumber)
        }
      } else {
        missing.push(docketNumber);
      }
    }
  });

  return { misplaced, missing, repairedData };
}

const sendToFilemaker = (data) => {
  const scriptName = "json * callback";
  const scriptParameter = JSON.stringify({ data });
  // Check if FileMaker object exists and call the script
  if (typeof window.FileMaker !== 'undefined') {
    window.FileMaker.PerformScript(scriptName, scriptParameter);
  } else {
    console.error("FileMaker object is not available.");
  }
}

window.loadApp = (json) => {
  const data = JSON.parse(json);
  console.log({ data });

  const rawData = data.rawData;
  const cleanedData = data.cleanedData;

  // Step 1: Extract docket numbers from rawData
  const docketNumbers = extractDocketNumbers(rawData);

  // Step 2: Check each docket number in the cleanedData for correct placement
  const result = checkDocketNumbersInCleanedData(docketNumbers, cleanedData);

  console.log(result); // For debugging

  // Step 3: Send the result to FileMaker
  sendToFilemaker(result);
}

window.convertToXLSX = (csvString) => {
  console.log("convert to XLSX called ...");

  try {
    // Parse the CSV data into a workbook
    const workbook = XLSX.read(csvString, { type: 'string' });

    // Generate XLSX data from the workbook
    const xlsxData = XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });

    sendToFilemaker(xlsxData);

  } catch (error) {
    console.error('Error occurred during XLSX conversion:', error.message);
    sendToFilemaker({
      status: 'failure',
      message: `Error occurred during XLSX conversion: ${error.message}`,
    });
  }
};

