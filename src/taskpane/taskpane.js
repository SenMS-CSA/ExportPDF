/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
import { saveAs } from 'file-saver'
import { PDFDocument } from 'pdf-lib'

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
   Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("exportPage").onclick = exportSpecificPage;
    document.getElementById("exportAllPages").onchange = togglePageNumberInput;
    populateSheetDropdown();
    getDocumentAsPDF();
  }
});

function togglePageNumberInput() {
  const checkbox = document.getElementById("exportAllPages");
  const pageNumberGroup = document.getElementById("pageNumberGroup");
  
  if (checkbox.checked) {
    pageNumberGroup.classList.add("disabled");
  } else {
    pageNumberGroup.classList.remove("disabled");
  }
}

async function populateSheetDropdown() {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const sheetSelect = document.getElementById("sheetSelect");
      sheetSelect.innerHTML = ""; // Clear loading message

      sheets.items.forEach((sheet) => {
        const option = document.createElement("option");
        option.value = sheet.name;
        option.textContent = sheet.name;
        sheetSelect.appendChild(option);
      });
    });
  } catch (error) {
    console.error("Error loading sheets:", error);
    const sheetSelect = document.getElementById("sheetSelect");
    sheetSelect.innerHTML = "<option value=''>Error loading sheets</option>";
  }
}

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      // Office.context.ui.displayDialogAsync('https://www.bing.com', { height: 30, width: 20, displayInIframe: true });
      console.log("Hello from Excel!");
     // getDocumentAsPDF();
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
function getDocumentAsPDF() {
  Office.context.document.getFileAsync(Office.FileType.Pdf,
    function (result) {
      if (result.status === "succeeded") {
        console.log("getFileAsync result: " + JSON.stringify(result.status));
        const myFile = result.value;
        const sliceCount = myFile.sliceCount;
        console.log("File size:" + myFile.size + " #Slices: " + sliceCount);

        // Get the file slices.
        const docDataSlices = [];
        let slicesReceived = 0, gotAllSlices = true;
        getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docDataSlices, slicesReceived, function(allSlices) {
          // This callback is called when all slices are collected
          createBlobFromSlices(allSlices, myFile);
        });
      } else {
        console.error("Failed to get file:", result.error);
      }
    }
  );
}

function createBlobFromSlices(docDataSlices, myFile) {
  // Flatten the array of arrays into a single array
  const flattenedData = docDataSlices.flat();
  const byteData = new Uint8Array(flattenedData);
  console.log("byteData length: " + byteData.length);
  const blob = new Blob([byteData], { type: 'application/pdf' });
  
  const blobUrl = URL.createObjectURL(blob);
  console.log("blobURL:" + blobUrl.toString());
  
  // Now you can use the blob/blobUrl
 Office.context.ui.openBrowserWindow(blobUrl);
//  printPDF(blobUrl);
  
  // Download the file
  const link = document.createElement('a');
  link.href = blobUrl;
  link.download = 'document.pdf';
  link.click();
  
  myFile.closeAsync();
}
function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docDataSlices, slicesReceived, onComplete) {
  file.getSliceAsync(nextSlice, function (sliceResult) {
    if (sliceResult.status == "succeeded") {
      if (!gotAllSlices) { /* Failed to get all slices, no need to continue. */
        return;
      }

      // Got one slice, store it in a temporary array.
      docDataSlices[sliceResult.value.index] = sliceResult.value.data;
      if (++slicesReceived == sliceCount) {
        // All slices have been received.
        onComplete(docDataSlices);
      }
      else {
        getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docDataSlices, slicesReceived, onComplete);
      }
    }
    else {
      gotAllSlices = false;
      file.closeAsync();
      console.log("getSliceAsync Error:", sliceResult.error.message);
    }
  });
}

function onGotAllSlices(docDataSlices) {
  let docData = [];
  for (let i = 0; i < docDataSlices.length; i++) {
    docData = docData.concat(docDataSlices[i]);
  }

  let fileContent = new String();
  for (let j = 0; j < docData.length; j++) {
    fileContent += String.fromCharCode(docData[j]);
  }

  // Now all the file content is stored in 'fileContent' variable,
  // you can do something with it, such as print, fax...
}
function printPDF(pdfUrl) {
    let win = window.open(pdfUrl, '_blank');
    win.focus();
    win.onload = function() {
        win.print();
    };
}

async function exportSpecificPage() {
  const pageNumberInput = document.getElementById("pageNumber");
  const sheetSelect = document.getElementById("sheetSelect");
  const statusElement = document.getElementById("exportStatus");
  const exportAllCheckbox = document.getElementById("exportAllPages");
  const exportAll = exportAllCheckbox.checked;
  const pageNumber = parseInt(pageNumberInput.value, 10);
  const selectedSheet = sheetSelect.value;

  if (!selectedSheet) {
    statusElement.textContent = "Please select a sheet.";
    statusElement.style.color = "red";
    return;
  }

  if (!exportAll && (!pageNumber || pageNumber < 1)) {
    statusElement.textContent = "Please enter a valid page number.";
    statusElement.style.color = "red";
    return;
  }

  if (exportAll) {
    statusElement.textContent = "Exporting all pages from '" + selectedSheet + "'...";
  } else {
    statusElement.textContent = "Exporting page " + pageNumber + " from '" + selectedSheet + "'...";
  }
  statusElement.style.color = "#666";
  statusElement.style.color = "#666";

  let hiddenSheets = [];

  try {
    // Hide all other sheets temporarily so only selected sheet is exported
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name, items/visibility");
      await context.sync();

      // Store original visibility and hide other sheets
      for (const sheet of sheets.items) {
        if (sheet.name !== selectedSheet) {
          if (sheet.visibility === Excel.SheetVisibility.visible) {
            hiddenSheets.push(sheet.name);
            sheet.visibility = Excel.SheetVisibility.hidden;
          }
        }
      }

      // Activate the selected sheet
      const targetSheet = context.workbook.worksheets.getItem(selectedSheet);
      targetSheet.activate();
      await context.sync();
    });

    // Small delay to ensure changes are applied before export
    await new Promise(resolve => setTimeout(resolve, 500));

    // Get the document as PDF (this will export only the visible sheet)
    Office.context.document.getFileAsync(Office.FileType.Pdf, async function(result) {
      if (result.status === "succeeded") {
        const myFile = result.value;
        const sliceCount = myFile.sliceCount;

        // Collect all slices
        const docDataSlices = [];
        let slicesReceived = 0;
        let gotAllSlices = true;

        collectSlices(myFile, 0, sliceCount, gotAllSlices, docDataSlices, slicesReceived, async function(allSlices) {
          try {
            // Flatten the array of arrays into a single array
            const flattenedData = allSlices.flat();
            const pdfBytes = new Uint8Array(flattenedData);

            // Load the PDF document
            const pdfDoc = await PDFDocument.load(pdfBytes);
            const totalPages = pdfDoc.getPageCount();

            let blob, downloadFileName, successMessage;

            if (exportAll) {
              // Export all pages - use the original PDF directly
              blob = new Blob([pdfBytes], { type: 'application/pdf' });
              downloadFileName = selectedSheet + '_all_pages.pdf';
              successMessage = "All " + totalPages + " page(s) from '" + selectedSheet + "' exported successfully!";
            } else {
              // Export specific page
              if (pageNumber > totalPages) {
                statusElement.textContent = "Page " + pageNumber + " does not exist. Sheet '" + selectedSheet + "' has " + totalPages + " page(s).";
                statusElement.style.color = "red";
                myFile.closeAsync();
                await restoreHiddenSheets(hiddenSheets);
                return;
              }

              // Create a new PDF with only the specified page
              const newPdfDoc = await PDFDocument.create();
              const [copiedPage] = await newPdfDoc.copyPages(pdfDoc, [pageNumber - 1]); // 0-indexed
              newPdfDoc.addPage(copiedPage);

              // Save the new PDF
              const newPdfBytes = await newPdfDoc.save();
              blob = new Blob([newPdfBytes], { type: 'application/pdf' });
              downloadFileName = selectedSheet + '_page_' + pageNumber + '.pdf';
              successMessage = "Page " + pageNumber + " from '" + selectedSheet + "' exported successfully!";
            }

            // Download the file
            const blobUrl = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = blobUrl;
            link.download = downloadFileName;
            link.click();

            // Open in browser
            Office.context.ui.openBrowserWindow(blobUrl);

            statusElement.textContent = successMessage;
            statusElement.style.color = "green";

            myFile.closeAsync();
            
            // Restore hidden sheets
            await restoreHiddenSheets(hiddenSheets);
          } catch (error) {
            console.error("Error processing PDF:", error);
            statusElement.textContent = "Error processing PDF: " + error.message;
            statusElement.style.color = "red";
            myFile.closeAsync();
            await restoreHiddenSheets(hiddenSheets);
          }
        });
      } else {
        statusElement.textContent = "Failed to get document: " + result.error.message;
        statusElement.style.color = "red";
        await restoreHiddenSheets(hiddenSheets);
      }
    });
  } catch (error) {
    console.error("Error:", error);
    statusElement.textContent = "Error: " + error.message;
    statusElement.style.color = "red";
    await restoreHiddenSheets(hiddenSheets);
  }
}

async function restoreHiddenSheets(hiddenSheets) {
  if (hiddenSheets.length === 0) return;
  
  try {
    await Excel.run(async (context) => {
      for (const sheetName of hiddenSheets) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        sheet.visibility = Excel.SheetVisibility.visible;
      }
      await context.sync();
    });
  } catch (error) {
    console.error("Error restoring sheets:", error);
  }
}

function collectSlices(file, nextSlice, sliceCount, gotAllSlices, docDataSlices, slicesReceived, onComplete) {
  file.getSliceAsync(nextSlice, function(sliceResult) {
    if (sliceResult.status === "succeeded") {
      if (!gotAllSlices) {
        return;
      }

      docDataSlices[sliceResult.value.index] = sliceResult.value.data;
      if (++slicesReceived === sliceCount) {
        onComplete(docDataSlices);
      } else {
        collectSlices(file, ++nextSlice, sliceCount, gotAllSlices, docDataSlices, slicesReceived, onComplete);
      }
    } else {
      gotAllSlices = false;
      file.closeAsync();
      console.error("getSliceAsync Error:", sliceResult.error.message);
    }
  });
}