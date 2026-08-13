import * as pdfjsLib from
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";


let zoomLevel = 100;

let pdfBlob = null;
let pdfURL = null;
let pdfName = "Converted Document.pdf";
let pdfDocument = null;


const pdfViewer =
  document.getElementById("pdfViewer");

const zoomLevelText =
  document.getElementById("zoomLevel");

const zoomInButton =
  document.getElementById("zoomIn");

const zoomOutButton =
  document.getElementById("zoomOut");

const printButton =
  document.getElementById("printButton");

const downloadButton =
  document.getElementById("downloadButton");

const emailButton =
  document.getElementById("emailButton");

const mobilePrint =
  document.getElementById("mobilePrint");

const mobileDownload =
  document.getElementById("mobileDownload");

const mobileEmail =
  document.getElementById("mobileEmail");

const documentName =
  document.getElementById("documentName");


// --------------------------------
// DATABASE
// --------------------------------

function openDatabase() {

  return new Promise((resolve, reject) => {

    const request =
      indexedDB.open("WordPDFEditor", 1);


    request.onupgradeneeded = () => {

      const db = request.result;

      if (
        !db.objectStoreNames.contains("documents")
      ) {

        db.createObjectStore("documents");
      }

    };


    request.onsuccess = () => {

      resolve(request.result);

    };


    request.onerror = () => {

      reject(request.error);

    };

  });

}


// --------------------------------
// GET SAVED PDF
// --------------------------------

async function getSavedPdf() {

  const db = await openDatabase();


  return new Promise((resolve, reject) => {

    const transaction =
      db.transaction(
        "documents",
        "readonly"
      );


    const store =
      transaction.objectStore(
        "documents"
      );


    const request =
      store.get("currentPdf");


    request.onsuccess = () => {

      resolve(request.result);

    };


    request.onerror = () => {

      reject(request.error);

    };

  });

}


// --------------------------------
// LOAD PDF
// --------------------------------

async function loadPdf() {

  const savedDocument =
    await getSavedPdf();


  if (!savedDocument) {

    showNoPdf();

    return;
  }


  pdfBlob =
    savedDocument.blob;

  pdfName =
    savedDocument.name ||
    "Converted Document.pdf";


  documentName.textContent =
    pdfName;


  pdfURL =
    URL.createObjectURL(pdfBlob);


  const pdfBytes =
    await pdfBlob.arrayBuffer();


  pdfDocument =
    await pdfjsLib.getDocument({
      data: pdfBytes
    }).promise;


  await renderPdf();

}


// --------------------------------
// RENDER ALL PDF PAGES
// --------------------------------

async function renderPdf() {

  if (!pdfDocument) {
    return;
  }


  pdfViewer.innerHTML = "";


  const scale =
    zoomLevel / 100;


  for (
    let pageNumber = 1;
    pageNumber <= pdfDocument.numPages;
    pageNumber++
  ) {

    const page =
      await pdfDocument.getPage(pageNumber);


    const viewport =
      page.getViewport({
        scale: scale
      });


    const pageWrapper =
      document.createElement("div");


    pageWrapper.className =
      "pdf-page";


    const canvas =
      document.createElement("canvas");


    const context =
      canvas.getContext("2d");


    const pixelRatio =
      window.devicePixelRatio || 1;


    canvas.width =
      Math.floor(
        viewport.width * pixelRatio
      );


    canvas.height =
      Math.floor(
        viewport.height * pixelRatio
      );


    canvas.style.width =
      `${viewport.width}px`;


    canvas.style.height =
      `${viewport.height}px`;


    pageWrapper.style.width =
      `${viewport.width}px`;


    pageWrapper.style.height =
      `${viewport.height}px`;


    pageWrapper.appendChild(canvas);


    pdfViewer.appendChild(
      pageWrapper
    );


    await page.render({

      canvasContext: context,

      viewport: viewport,

      transform:
        pixelRatio !== 1
          ? [
              pixelRatio,
              0,
              0,
              pixelRatio,
              0,
              0
            ]
          : null

    }).promise;

  }

}


// --------------------------------
// NO PDF
// --------------------------------

function showNoPdf() {

  pdfViewer.innerHTML = `

    <div class="empty-preview">

      <div class="document-icon">
        PDF
      </div>

      <h2>No PDF loaded</h2>

      <p>
        Return to WordPDF and convert
        a Word document first.
      </p>

    </div>

  `;

}


// --------------------------------
// ZOOM
// --------------------------------

async function updateZoom() {

  zoomLevelText.textContent =
    `${zoomLevel}%`;


  if (pdfDocument) {

    await renderPdf();

  }

}


zoomInButton.addEventListener(
  "click",
  async () => {

    if (zoomLevel < 200) {

      zoomLevel += 10;

      await updateZoom();

    }

  }
);


zoomOutButton.addEventListener(
  "click",
  async () => {

    if (zoomLevel > 50) {

      zoomLevel -= 10;

      await updateZoom();

    }

  }
);


// --------------------------------
// DOWNLOAD
// --------------------------------

function downloadDocument() {

  if (!pdfBlob || !pdfURL) {

    alert(
      "No PDF is currently loaded."
    );

    return;
  }


  const link =
    document.createElement("a");


  link.href = pdfURL;

  link.download = pdfName;


  document.body.appendChild(link);

  link.click();

  link.remove();

}


downloadButton.addEventListener(
  "click",
  downloadDocument
);


mobileDownload.addEventListener(
  "click",
  downloadDocument
);


// --------------------------------
// PRINT
// --------------------------------

function printDocument() {

  if (!pdfURL) {

    alert(
      "No PDF is currently loaded."
    );

    return;
  }


  const printFrame =
    document.createElement("iframe");


  printFrame.style.position =
    "fixed";

  printFrame.style.right =
    "0";

  printFrame.style.bottom =
    "0";

  printFrame.style.width =
    "0";

  printFrame.style.height =
    "0";

  printFrame.style.border =
    "0";


  printFrame.src =
    pdfURL;


  document.body.appendChild(
    printFrame
  );


  printFrame.onload = () => {

    setTimeout(() => {

      try {

        printFrame.contentWindow.focus();

        printFrame.contentWindow.print();

      } catch (error) {

        console.error(error);

        window.open(
          pdfURL,
          "_blank"
        );

      }

    }, 500);

  };

}


printButton.addEventListener(
  "click",
  printDocument
);


mobilePrint.addEventListener(
  "click",
  printDocument
);


// --------------------------------
// EMAIL / SHARE
// --------------------------------

async function emailDocument() {

  if (!pdfBlob) {

    alert(
      "No PDF is currently loaded."
    );

    return;
  }


  const file =
    new File(
      [pdfBlob],
      pdfName,
      {
        type: "application/pdf"
      }
    );


  try {

    if (
      navigator.share &&
      navigator.canShare &&
      navigator.canShare({
        files: [file]
      })
    ) {

      await navigator.share({

        title: pdfName,

        text:
          "Converted with WordPDF",

        files: [file]

      });


      return;

    }

  } catch (error) {

    if (
      error.name === "AbortError"
    ) {

      return;

    }


    console.error(error);

  }


  alert(
    "Your browser cannot share the PDF directly. Download the file and attach it to your email."
  );

}


emailButton.addEventListener(
  "click",
  emailDocument
);


mobileEmail.addEventListener(
  "click",
  emailDocument
);


// --------------------------------
// EDITING TOOL SELECTION
// --------------------------------

const toolButtons =
  document.querySelectorAll(
    ".tool-button"
  );


toolButtons.forEach((button) => {

  button.addEventListener(
    "click",
    () => {

      toolButtons.forEach(
        (item) => {

          item.classList.remove(
            "active-tool"
          );

        }
      );


      button.classList.add(
        "active-tool"
      );

    }
  );

});


// --------------------------------
// START
// --------------------------------

zoomLevelText.textContent =
  `${zoomLevel}%`;


loadPdf().catch((error) => {

  console.error(error);


  pdfViewer.innerHTML = `

    <div class="empty-preview">

      <div class="document-icon">
        PDF
      </div>

      <h2>Unable to display PDF</h2>

      <p>
        Return to WordPDF and convert
        the document again.
      </p>

    </div>

  `;

});
