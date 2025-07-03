// ===============================================================
// GATE PASS MANAGEMENT SYSTEM - SERVER-SIDE LOGIC (Code.gs)
// ===============================================================

// --- CONFIGURATION ---
const SPREADSHEET_ID = "1Yj7ksiSyeGUVBrqx4U7fzri5l-h5zKwS5qd7zYi_1Rw";
const SHEET_NAME = "GatePassLog";
const USERS_SHEET_NAME = "Users";
const DRIVE_FOLDER_ID = "1rFPLMPNzytP2MT9q_1_MYrqfdEmvjMIw";

/**
 * Main function to serve the HTML interface
 */
function doGet() {
  const html = HtmlService.createTemplateFromFile('index').evaluate();
  html.setTitle('Hotel Gate Pass System');
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  return html;
}

/**
 * Gets the main spreadsheet object
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Gets the gate pass log sheet
 */
function getLogSheet() {
  return getSpreadsheet().getSheetByName(SHEET_NAME);
}

/**
 * Gets the users sheet
 */
function getUsersSheet() {
  let spreadsheet = getSpreadsheet();
  let sheet = spreadsheet.getSheetByName(USERS_SHEET_NAME);
  
  // Create users sheet if it doesn't exist
  if (!sheet) {
    sheet = spreadsheet.insertSheet(USERS_SHEET_NAME);
    sheet.appendRow(['Email', 'Name', 'Role', 'Department', 'Created']);
  }
  
  return sheet;
}

/**
 * Submits a new gate pass to the system
 */
function submitGatePass({formData, fileData}) {
  try {
    const sheet = getLogSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Generate new ID and prepare dates
    const newId = generateNewPassId(formData.passType);
    const issueDate = new Date(formData.issueDate);
    const expectedReturnDate = formData.expectedReturnDate ? new Date(formData.expectedReturnDate) : '';
    
    // Handle image upload if present
    let imageUrl = '';
    if (fileData) {
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const fileName = `${newId}_${fileData.fileName}`;
      const blob = Utilities.newBlob(Utilities.base64Decode(fileData.bytes), fileData.mimeType, fileName);
      const file = folder.createFile(blob);
      imageUrl = file.getUrl();
    }

    // Prepare the new row data in the correct column order
    const newRow = [];
    headers.forEach(header => {
      switch(header) {
        case 'GatePassID':
          newRow.push(newId);
          break;
        case 'PassType':
          newRow.push(formData.passType);
          break;
        case 'IssueDate':
          newRow.push(issueDate);
          break;
        case 'IssuedTo':
          newRow.push(formData.issuedTo);
          break;
        case 'Department':
          newRow.push(formData.department);
          break;
        case 'ItemsJSON':
          newRow.push(JSON.stringify(formData.items));
          break;
        case 'Purpose':
          newRow.push(formData.purpose);
          break;
        case 'AuthorizedBy':
          newRow.push(formData.authorizedBy);
          break;
        case 'Status':
          newRow.push(formData.passType === 'RGP' ? 'Pending Return' : 'Issued');
          break;
        case 'ExpectedReturnDate':
          newRow.push(expectedReturnDate);
          break;
        case 'ActualReturnDate':
          newRow.push('');
          break;
        case 'CreatedTimestamp':
          newRow.push(new Date());
          break;
        case 'Remarks':
          newRow.push(formData.remarks || '');
          break;
        case 'ImageURL':
          newRow.push(imageUrl || '');
          break;
        default:
          newRow.push(''); // For any unexpected columns
      }
    });

    sheet.appendRow(newRow);

    return { 
      status: 'success', 
      message: `Gate Pass ${newId} created successfully!`,
      newPassId: newId 
    };

  } catch (error) {
    console.error(error);
    return { 
      status: 'error', 
      message: 'An error occurred: ' + error.message 
    };
  }
}

/**
 * Generates a new unique Gate Pass ID
 */
function generateNewPassId(passType) {
  const sheet = getLogSheet();
  const lastRow = sheet.getLastRow();
  const year = new Date().getFullYear();
  let newNumber = 1;

  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const passIdsOfType = data
      .map(row => row[0])
      .filter(id => id && id.startsWith(passType + '-' + year));

    if (passIdsOfType.length > 0) {
      const lastId = passIdsOfType[passIdsOfType.length - 1];
      const lastNumber = parseInt(lastId.split('-')[2], 10) || 0;
      newNumber = lastNumber + 1;
    }
  }

  const formattedNumber = ('000' + newNumber).slice(-3);
  return `${passType}-${year}-${formattedNumber}`;
}

/**
 * Gets all pending returnable passes (RGP with no actual return date)
 */
function getPendingReturnablePasses() {
  try {
    const sheet = getLogSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    const data = values.slice(1);

    const pendingPasses = data.map((row, index) => {
      const pass = {};
      headers.forEach((header, i) => {
        pass[header] = row[i];
      });
      pass.rowNumber = index + 2;
      
      // Parse items JSON if exists
      if (pass.ItemsJSON) {
        try {
          pass.ItemsJSON = JSON.parse(pass.ItemsJSON);
        } catch (e) {
          pass.ItemsJSON = [];
        }
      } else {
        pass.ItemsJSON = [];
      }
      
      return pass;
    }).filter(pass => pass.PassType === 'RGP' && pass.Status !== 'Returned');
    
    // Format dates for display
    pendingPasses.forEach(pass => {
      if (pass.IssueDate instanceof Date) {
        pass.IssueDate = pass.IssueDate.toLocaleDateString();
      }
      if (pass.ExpectedReturnDate instanceof Date) {
        pass.ExpectedReturnDate = pass.ExpectedReturnDate.toLocaleDateString();
      }
    });

    return { status: 'success', data: pendingPasses };
  } catch (error) {
    console.error(error);
    return { 
      status: 'error', 
      message: 'Could not fetch pending passes: ' + error.message, 
      data: [] 
    };
  }
}

/**
 * Gets all overdue returnable passes (RGP with expected return date in the past)
 */
function getOverduePasses() {
  try {
    const sheet = getLogSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    const data = values.slice(1);

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const overduePasses = data.map((row, index) => {
      const pass = {};
      headers.forEach((header, i) => {
        pass[header] = row[i];
      });
      pass.rowNumber = index + 2;
      
      // Parse items JSON if exists
      if (pass.ItemsJSON) {
        try {
          pass.ItemsJSON = JSON.parse(pass.ItemsJSON);
        } catch (e) {
          pass.ItemsJSON = [];
        }
      } else {
        pass.ItemsJSON = [];
      }
      
      return pass;
    }).filter(pass => {
      if (pass.PassType !== 'RGP' || pass.Status === 'Returned') return false;
      
      if (pass.ExpectedReturnDate instanceof Date) {
        const returnDate = new Date(pass.ExpectedReturnDate);
        returnDate.setHours(0, 0, 0, 0);
        return returnDate < today;
      }
      return false;
    });
    
    // Format dates for display
    overduePasses.forEach(pass => {
      if (pass.IssueDate instanceof Date) {
        pass.IssueDate = pass.IssueDate.toLocaleDateString();
      }
      if (pass.ExpectedReturnDate instanceof Date) {
        pass.ExpectedReturnDate = pass.ExpectedReturnDate.toLocaleDateString();
      }
    });

    return { status: 'success', data: overduePasses };
  } catch (error) {
    console.error(error);
    return { 
      status: 'error', 
      message: 'Could not fetch overdue passes: ' + error.message, 
      data: [] 
    };
  }
}

/**
 * Marks a pass as returned with partial/full returns
 */
function markAsReturned({rowNumber, passId, actualReturnDate, returnRemarks, returnedItems, pendingItems}) {
  try {
    const sheet = getLogSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const statusCol = headers.indexOf('Status') + 1;
    const actualReturnDateCol = headers.indexOf('ActualReturnDate') + 1;
    const remarksCol = headers.indexOf('Remarks') + 1;
    const itemsJsonCol = headers.indexOf('ItemsJSON') + 1;
    const idCol = headers.indexOf('GatePassID') + 1;

    if (statusCol === 0 || actualReturnDateCol === 0 || idCol === 0) {
      throw new Error("Could not find required columns in the sheet.");
    }
    
    // Verify we're updating the correct row
    const currentId = sheet.getRange(rowNumber, idCol).getValue();
    if (currentId !== passId) {
      return { 
        status: 'error', 
        message: `ID mismatch. Expected ${passId} but found ${currentId}.` 
      };
    }

    // Update the status and return date
    sheet.getRange(rowNumber, statusCol).setValue('Returned');
    sheet.getRange(rowNumber, actualReturnDateCol).setValue(new Date(actualReturnDate));
    
    // Update remarks if provided
    if (returnRemarks) {
      const currentRemarks = sheet.getRange(rowNumber, remarksCol).getValue();
      sheet.getRange(rowNumber, remarksCol).setValue(
        currentRemarks ? `${currentRemarks}\nReturn Remarks: ${returnRemarks}` : `Return Remarks: ${returnRemarks}`
      );
    }
    
    // If there are pending items, create a new RGP for them
    if (pendingItems && pendingItems.length > 0) {
      const originalPass = getPassData(rowNumber);
      
      const newId = generateNewPassId('RGP');
      const newRow = [];
      headers.forEach(header => {
        switch(header) {
          case 'GatePassID':
            newRow.push(newId);
            break;
          case 'PassType':
            newRow.push('RGP');
            break;
          case 'IssueDate':
            newRow.push(new Date());
            break;
          case 'IssuedTo':
            newRow.push(originalPass.IssuedTo);
            break;
          case 'Department':
            newRow.push(originalPass.Department);
            break;
          case 'ItemsJSON':
            newRow.push(JSON.stringify(pendingItems));
            break;
          case 'Purpose':
            newRow.push(`Pending items from ${passId}: ${originalPass.Purpose}`);
            break;
          case 'AuthorizedBy':
            newRow.push(originalPass.AuthorizedBy);
            break;
          case 'Status':
            newRow.push('Pending Return');
            break;
          case 'ExpectedReturnDate':
            newRow.push('');
            break;
          case 'ActualReturnDate':
            newRow.push('');
            break;
          case 'CreatedTimestamp':
            newRow.push(new Date());
            break;
          case 'Remarks':
            newRow.push(`Pending items from ${passId}: ${originalPass.Remarks || ''}`);
            break;
          case 'ImageURL':
            newRow.push(originalPass.ImageURL || '');
            break;
          default:
            newRow.push('');
        }
      });
      
      sheet.appendRow(newRow);
    }

    return { 
      status: 'success', 
      message: pendingItems && pendingItems.length > 0 ? 
        `Pass ${passId} partially returned. New RGP created for pending items.` : 
        `Pass ${passId} has been fully returned.` 
    };
  } catch (error) {
    console.error(error);
    return { 
      status: 'error', 
      message: 'An error occurred while updating: ' + error.message 
    };
  }
}

/**
 * Gets pass data for a specific row
 */
function getPassData(rowNumber) {
  const sheet = getLogSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const pass = {};
  headers.forEach((header, i) => {
    pass[header] = rowData[i];
  });
  
  if (pass.ItemsJSON) {
    try {
      pass.ItemsJSON = JSON.parse(pass.ItemsJSON);
    } catch (e) {
      pass.ItemsJSON = [];
    }
  } else {
    pass.ItemsJSON = [];
  }
  
  return pass;
}

/**
 * Updates an existing gate pass
 */
function updateGatePass({rowNumber, passId, formData}) {
  try {
    const sheet = getLogSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Verify we're updating the correct row
    const currentId = sheet.getRange(rowNumber, headers.indexOf('GatePassID') + 1).getValue();
    if (currentId !== passId) {
      return { 
        status: 'error', 
        message: `ID mismatch. Expected ${passId} but found ${currentId}.` 
      };
    }

    // Update each field
    headers.forEach(header => {
      const col = headers.indexOf(header) + 1;
      switch(header) {
        case 'IssuedTo':
          sheet.getRange(rowNumber, col).setValue(formData.issuedTo);
          break;
        case 'Department':
          sheet.getRange(rowNumber, col).setValue(formData.department);
          break;
        case 'IssueDate':
          sheet.getRange(rowNumber, col).setValue(new Date(formData.issueDate));
          break;
        case 'ExpectedReturnDate':
          sheet.getRange(rowNumber, col).setValue(formData.expectedReturnDate ? new Date(formData.expectedReturnDate) : '');
          break;
        case 'Purpose':
          sheet.getRange(rowNumber, col).setValue(formData.purpose);
          break;
        case 'AuthorizedBy':
          sheet.getRange(rowNumber, col).setValue(formData.authorizedBy);
          break;
        case 'Remarks':
          sheet.getRange(rowNumber, col).setValue(formData.remarks || '');
          break;
        case 'ItemsJSON':
          sheet.getRange(rowNumber, col).setValue(JSON.stringify(formData.items));
          break;
      }
    });

    return { 
      status: 'success', 
      message: `Pass ${passId} has been updated successfully.` 
    };
  } catch (error) {
    console.error(error);
    return { 
      status: 'error', 
      message: 'An error occurred while updating: ' + error.message 
    };
  }
}

/**
 * Gets all users from the system
 */
function getUsers() {
  try {
    const sheet = getUsersSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // Skip header row
    const data = values.slice(1);
    
    const users = data.map(row => ({
      email: row[0],
      name: row[1],
      role: row[2],
      department: row[3],
      created: row[4]
    }));
    
    return { status: 'success', data: users };
  } catch (error) {
    console.error(error);
    return { 
      status: 'error', 
      message: 'Could not fetch users: ' + error.message, 
      data: [] 
    };
  }
}

/**
 * Adds a new user to the system
 */
function addUser(userData) {
  try {
    const sheet = getUsersSheet();
    
    // Check if user already exists
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const emails = values.slice(1).map(row => row[0]);
    
    if (emails.includes(userData.email)) {
      return { 
        status: 'error', 
        message: 'User with this email already exists.' 
      };
    }
    
    // Add new user
    sheet.appendRow([
      userData.email,
      userData.name,
      userData.role,
      userData.department,
      new Date()
    ]);
    
    return { 
      status: 'success', 
      message: `User ${userData.email} added successfully.` 
    };
  } catch (error) {
    console.error(error);
    return { 
      status: 'error', 
      message: 'An error occurred while adding user: ' + error.message 
    };
  }
}

/**
 * Deletes a user from the system
 */
function deleteUser(email) {
  try {
    const sheet = getUsersSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // Find user row
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === email) {
        sheet.deleteRow(i + 1);
        return { 
          status: 'success', 
          message: `User ${email} deleted successfully.` 
        };
      }
    }
    
    return { 
      status: 'error', 
      message: 'User not found.' 
    };
  } catch (error) {
    console.error(error);
    return { 
      status: 'error', 
      message: 'An error occurred while deleting user: ' + error.message 
    };
  }
}

/**
 * Generates a PDF for a specific gate pass
 */
function generateGatePassPDF(passId) {
  try {
    const sheet = getLogSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    
    // Find the pass
    for (let i = 1; i < values.length; i++) {
      if (values[i][headers.indexOf('GatePassID')] === passId) {
        const pass = {};
        headers.forEach((header, index) => {
          pass[header] = values[i][index];
        });
        
        // Parse items JSON if exists
        if (pass.ItemsJSON) {
          try {
            pass.ItemsJSON = JSON.parse(pass.ItemsJSON);
          } catch (e) {
            pass.ItemsJSON = [];
          }
        } else {
          pass.ItemsJSON = [];
        }
        
        // Format dates
        if (pass.IssueDate instanceof Date) {
          pass.IssueDate = pass.IssueDate.toLocaleDateString();
        }
        if (pass.ExpectedReturnDate instanceof Date) {
          pass.ExpectedReturnDate = pass.ExpectedReturnDate.toLocaleDateString();
        }
        
        // Generate PDF (this is simplified - in a real app you would use a proper PDF library)
        const pdfContent = `
          <html>
            <body>
              <h1>${pass.PassType === 'RGP' ? 'RETURNABLE GATE PASS (RGP)' : 'NON-RETURNABLE GATE PASS (NRGP)'}</h1>
              <p>Pass ID: ${pass.GatePassID}</p>
              <hr>
              <p>Issued To: ${pass.IssuedTo}</p>
              <p>Department: ${pass.Department}</p>
              <p>Issue Date: ${pass.IssueDate}</p>
              <p>Expected Return: ${pass.PassType === 'RGP' ? pass.ExpectedReturnDate : 'N/A'}</p>
              <p>Purpose: ${pass.Purpose}</p>
              <p>Authorized By: ${pass.AuthorizedBy}</p>
              <hr>
              <h2>Items:</h2>
              <ul>
                ${pass.ItemsJSON.map(item => `<li>${item.quantity} ${item.unit || ''} - ${item.description}</li>`).join('')}
              </ul>
              ${pass.Remarks ? `<p>Remarks: ${pass.Remarks}</p>` : ''}
              <hr>
              <p>Generated on: ${new Date().toLocaleDateString()}</p>
            </body>
          </html>
        `;
        
        return { 
          status: 'success', 
          message: 'PDF generated successfully.',
          pdfContent: pdfContent 
        };
      }
    }
    
    return { 
      status: 'error', 
      message: 'Pass not found.' 
    };
  } catch (error) {
    console.error(error);
    return { 
      status: 'error', 
      message: 'An error occurred while generating PDF: ' + error.message 
    };
  }
}