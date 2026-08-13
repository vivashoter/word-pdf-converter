import * as pdfjsLib from
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";


let zoomLevel = 100;

let pdfBlob = null;
let pdfURL = null;
let pdfName = "Converted Document.pdf";
let pdfDocument = null;

let activeTool = "select";

let isDrawing = false;
let isHighlighting = false;

let startX = 0;
let startY = 0;

let currentDrawCanvas = null;
let currentDrawContext = null;

let currentHighlight = null;


// --------------------------------
// ELEMENTS
// --------------------------------

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


async function getSavedPdf() {

  const db =
    await openDatabase();


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
// RENDER PDF
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
      await pdfDocument.getPage(
        pageNumber
      );


    const viewport =
      page.getViewport({
        scale: scale
      });


    const pageWrapper =
      document.createElement("div");


    pageWrapper.className =
      "pdf-page";


    pageWrapper.dataset.page =
      pageNumber;


    pageWrapper.style.width =
      `${viewport.width}px`;


    pageWrapper.style.height =
      `${viewport.height}px`;


    // PDF CANVAS

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


    canvas.className =
      "pdf-canvas";


    pageWrapper.appendChild(
      canvas
    );


    // ANNOTATION LAYER

    const annotationLayer =
      document.createElement("div");


    annotationLayer.className =
      "annotation-layer";


    pageWrapper.appendChild(
      annotationLayer
    );


    // DRAWING CANVAS

    const drawCanvas =
      document.createElement("canvas");


    drawCanvas.className =
      "draw-layer";


    drawCanvas.width =
      Math.floor(
        viewport.width * pixelRatio
      );


    drawCanvas.height =
      Math.floor(
        viewport.height * pixelRatio
      );


    drawCanvas.style.width =
      `${viewport.width}px`;


    drawCanvas.style.height =
      `${viewport.height}px`;


    drawCanvas.dataset.ratio =
      pixelRatio;


    pageWrapper.appendChild(
      drawCanvas
    );


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


    setupPageEditing(
      pageWrapper
    );

  }


  updateToolMode();

}


// --------------------------------
// PAGE EDITING
// --------------------------------

function setupPageEditing(page) {

  page.addEventListener(
    "pointerdown",
    handlePointerDown
  );


  page.addEventListener(
    "pointermove",
    handlePointerMove
  );


  page.addEventListener(
    "pointerup",
    handlePointerUp
  );


  page.addEventListener(
    "pointercancel",
    handlePointerUp
  );

}


// --------------------------------
// POINTER POSITION
// --------------------------------

function getPosition(event, page) {

  const rect =
    page.getBoundingClientRect();


  return {

    x:
      event.clientX -
      rect.left,

    y:
      event.clientY -
      rect.top

  };

}


// --------------------------------
// POINTER DOWN
// --------------------------------

function handlePointerDown(event) {

  const page =
    event.currentTarget;


  const position =
    getPosition(
      event,
      page
    );


  // TEXT

  if (activeTool === "text") {

    createText(
      page,
      position.x,
      position.y
    );

    return;
  }


  // DRAW

  if (activeTool === "draw") {

    event.preventDefault();

    isDrawing = true;


    currentDrawCanvas =
      page.querySelector(
        ".draw-layer"
      );


    currentDrawContext =
      currentDrawCanvas.getContext(
        "2d"
      );


    const ratio =
      Number(
        currentDrawCanvas.dataset.ratio
      ) || 1;


    currentDrawContext.lineWidth =
      3 * ratio;


    currentDrawContext.lineCap =
      "round";


    currentDrawContext.lineJoin =
      "round";


    currentDrawContext.strokeStyle =
      "#111827";


    currentDrawContext.beginPath();


    currentDrawContext.moveTo(
      position.x * ratio,
      position.y * ratio
    );


    page.setPointerCapture(
      event.pointerId
    );


    return;
  }


  // HIGHLIGHT

  if (
    activeTool === "highlight"
  ) {

    event.preventDefault();

    isHighlighting = true;


    startX =
      position.x;

    startY =
      position.y;


    const annotationLayer =
      page.querySelector(
        ".annotation-layer"
      );


    currentHighlight =
      document.createElement(
        "div"
      );


    currentHighlight.className =
      "highlight-annotation";


    currentHighlight.style.left =
      `${startX}px`;


    currentHighlight.style.top =
      `${startY}px`;


    annotationLayer.appendChild(
      currentHighlight
    );


    page.setPointerCapture(
      event.pointerId
    );

  }

}


// --------------------------------
// POINTER MOVE
// --------------------------------

function handlePointerMove(event) {

  const page =
    event.currentTarget;


  const position =
    getPosition(
      event,
      page
    );


  // DRAW

  if (
    activeTool === "draw" &&
    isDrawing &&
    currentDrawContext
  ) {

    event.preventDefault();


    const ratio =
      Number(
        currentDrawCanvas.dataset.ratio
      ) || 1;


    currentDrawContext.lineTo(
      position.x * ratio,
      position.y * ratio
    );


    currentDrawContext.stroke();


    return;
  }


  // HIGHLIGHT

  if (
    activeTool === "highlight" &&
    isHighlighting &&
    currentHighlight
  ) {

    event.preventDefault();


    const left =
      Math.min(
        startX,
        position.x
      );


    const top =
      Math.min(
        startY,
        position.y
      );


    const width =
      Math.abs(
        position.x -
        startX
      );


    const height =
      Math.abs(
        position.y -
        startY
      );


    currentHighlight.style.left =
      `${left}px`;


    currentHighlight.style.top =
      `${top}px`;


    currentHighlight.style.width =
      `${width}px`;


    currentHighlight.style.height =
      `${height}px`;

  }

}


// --------------------------------
// POINTER UP
// --------------------------------

function handlePointerUp(event) {

  if (isDrawing) {

    isDrawing = false;


    if (currentDrawContext) {

      currentDrawContext.closePath();

    }


    currentDrawCanvas = null;
    currentDrawContext = null;

  }


  if (isHighlighting) {

    isHighlighting = false;


    if (currentHighlight) {

      const width =
        parseFloat(
          currentHighlight.style.width
        );


      const height =
        parseFloat(
          currentHighlight.style.height
        );


      if (
        width < 5 ||
        height < 5
      ) {

        currentHighlight.remove();

      }

    }


    currentHighlight = null;

  }

}


// --------------------------------
// CREATE TEXT
// --------------------------------

function createText(
  page,
  x,
  y
) {

  const annotationLayer =
    page.querySelector(
      ".annotation-layer"
    );


  const text =
    document.createElement("div");


  text.className =
    "text-annotation";


  text.contentEditable =
    "true";


  text.textContent =
    "Type here";


  text.style.left =
    `${x}px`;


  text.style.top =
    `${y}px`;


  annotationLayer.appendChild(
    text
  );


  text.focus();


  const range =
    document.createRange();


  range.selectNodeContents(
    text
  );


  const selection =
    window.getSelection();


  selection.removeAllRanges();

  selection.addRange(range);

}


// --------------------------------
// TOOL MODE
// --------------------------------

function updateToolMode() {

  const pages =
    document.querySelectorAll(
      ".pdf-page"
    );


  pages.forEach((page) => {

    page.dataset.tool =
      activeTool;

  });

}


// --------------------------------
// TOOLBAR
// --------------------------------

const toolButtons =
  document.querySelectorAll(
    ".tool-button"
  );


toolButtons.forEach((button) => {

  button.addEventListener(
    "click",
    () => {

      const tool =
        button.dataset.tool;


      // Image and Sign come next

      if (
        tool === "image" ||
        tool === "sign"
      ) {

        alert(
          `${tool === "image" ? "Image" : "Sign"} will be added next.`
        );

        return;
      }


      activeTool =
        tool || "select";


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


      updateToolMode();

  });

});


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
// CURRENTLY ORIGINAL PDF
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


  link.href =
    pdfURL;


  link.download =
    pdfName;


  document.body.appendChild(
    link
  );


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


  window.open(
    pdfURL,
    "_blank"
  );

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
        type:
          "application/pdf"
      }
    );


  if (
    navigator.share &&
    navigator.canShare &&
    navigator.canShare({
      files: [file]
    })
  ) {

    try {

      await navigator.share({

        title: pdfName,

        text:
          "Converted with WordPDF",

        files:
          [file]

      });


      return;

    }

    catch (error) {

      if (
        error.name !==
        "AbortError"
      ) {

        console.error(error);

      }


      return;

    }

  }


  alert(
    "Your browser cannot share the PDF directly. Download it and attach it to your email."
  );

}


// --------------------------------
// EMAIL BUTTONS
// --------------------------------

emailButton.addEventListener(
  "click",
  emailDocument
);


mobileEmail.addEventListener(
  "click",
  emailDocument
);


// --------------------------------
// EMPTY STATE
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
        Return to WordPDF and
        convert the document again.
      </p>

    </div>

  `;

});
