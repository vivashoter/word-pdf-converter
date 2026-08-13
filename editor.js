let zoomLevel = 100;

let pdfBlob = null;
let pdfURL = null;
let pdfName = "Converted Document.pdf";


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


    request.onupgradeneeded =
      function () {

        const db =
          request.result;

        if (
          !db.objectStoreNames.contains(
            "documents"
          )
        ) {

          db.createObjectStore(
            "documents"
          );
        }

      };


    request.onsuccess =
      function () {

        resolve(request.result);

      };


    request.onerror =
      function () {

        reject(request.error);

      };

  });

}


async function loadPdf() {

  const db =
    await openDatabase();


  const savedDocument =
    await new Promise(
      (resolve, reject) => {

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
          store.get(
            "currentPdf"
          );


        request.onsuccess =
          function () {

            resolve(
              request.result
            );

          };


        request.onerror =
          function () {

            reject(
              request.error
            );

          };

      }
    );


  if (!savedDocument) {

    pdfViewer.innerHTML = `

      <div class="empty-preview">

        <div class="document-icon">
          PDF
        </div>

        <h2>No PDF loaded</h2>

        <p>
          Convert a Word document first
          to open it in the editor.
        </p>

      </div>

    `;

    return;
  }


  pdfBlob =
    savedDocument.blob;


  pdfName =
    savedDocument.name;


  documentName.textContent =
    pdfName;


  pdfURL =
    URL.createObjectURL(
      pdfBlob
    );


  displayPdf();

}


// --------------------------------
// DISPLAY PDF
// --------------------------------

function displayPdf() {

  pdfViewer.innerHTML = "";


  const iframe =
    document.createElement(
      "iframe"
    );


  iframe.src =
    pdfURL + "#toolbar=0";


  iframe.style.width =
    "100%";


  iframe.style.height =
    "1050px";


  iframe.style.border =
    "none";


  iframe.title =
    pdfName;


  pdfViewer.appendChild(
    iframe
  );

}


// --------------------------------
// ZOOM
// --------------------------------

function updateZoom() {

  pdfViewer.style.transform =
    `scale(${zoomLevel / 100})`;


  pdfViewer.style.transformOrigin =
    "top center";


  zoomLevelText.textContent =
    `${zoomLevel}%`;

}


zoomInButton.addEventListener(
  "click",
  () => {

    if (zoomLevel < 200) {

      zoomLevel += 10;

      updateZoom();

    }

  }
);


zoomOutButton.addEventListener(
  "click",
  () => {

    if (zoomLevel > 50) {

      zoomLevel -= 10;

      updateZoom();

    }

  }
);


// --------------------------------
// DOWNLOAD
// --------------------------------

function downloadDocument() {

  if (!pdfBlob) {

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

  if (!pdfBlob) {

    alert(
      "No PDF is currently loaded."
    );

    return;
  }


  const printWindow =
    window.open(
      pdfURL,
      "_blank"
    );


  if (!printWindow) {

    alert(
      "Please allow popups to print this PDF."
    );

    return;
  }


  printWindow.onload =
    function () {

      printWindow.print();

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


  if (
    navigator.share &&
    navigator.canShare &&
    navigator.canShare({
      files: [file]
    })
  ) {

    try {

      await navigator.share({

        title:
          "WordPDF Document",

        text:
          "Here is the converted PDF.",

        files:
          [file]

      });

      return;

    }

    catch (error) {

      if (
        error.name !== "AbortError"
      ) {

        console.error(error);

      }

      return;

    }

  }


  alert(
    "Direct file sharing is not supported by this browser yet. Download the PDF and attach it to your email."
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
// TOOLBAR
// --------------------------------

const toolButtons =
  document.querySelectorAll(
    ".tool-button"
  );


toolButtons.forEach(
  (button) => {

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

  }
);


// --------------------------------
// START
// --------------------------------

updateZoom();

loadPdf().catch(
  (error) => {

    console.error(error);

    pdfViewer.innerHTML = `

      <div class="empty-preview">

        <div class="document-icon">
          PDF
        </div>

        <h2>Unable to load PDF</h2>

        <p>
          Please return to WordPDF
          and convert the document again.
        </p>

      </div>

    `;

  }
);
