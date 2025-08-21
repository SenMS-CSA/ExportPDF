/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
import { saveAs } from 'file-saver'
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
   Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    getDocumentAsPDF();
  }
});

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