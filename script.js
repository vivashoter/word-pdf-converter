const wordToPdfButton = document.getElementById("wordToPdf");
const pdfToWordButton = document.getElementById("pdfToWord");

const fileInput = document.getElementById("fileInput");
const uploadTitle = document.getElementById("uploadTitle");
const uploadDescription = document.getElementById("uploadDescription");
const uploadArea = document.getElementById("uploadArea");

const convertButton = document.getElementById("convertButton");
const statusMessage = document.getElementById("statusMessage");

let conversionMode = "word-to-pdf";


// ---------------------------------
// SAVE PDF FOR EDITOR
// ---------------------------------

function openDatabase() {

  return new Promise((resolve, reject) => {

    const request =
      indexedDB.open("WordPDFEditor", 1);

    request.onupgradeneeded = function () {

      const db = request.result;

      if (!db.objectStoreNames.contains("documents")) {

        db.createObjectStore("documents");
      }
    };

    request.onsuccess = function () {

      resolve(request.result);
    };

    request.onerror = function () {

      reject(request.error);
    };

  });
}


async function savePdfForEditor(blob, fileName) {

  const db = await openDatabase();

  return new Promise((resolve, reject) => {

    const transaction =
      db.transaction(
        "documents",
        "readwrite"
      );

    const store =
      transaction.objectStore(
        "documents"
      );

    store.put(
      {
        blob: blob,
        name: fileName
      },
      "currentPdf"
    );

    transaction.oncomplete = function () {

      resolve();
    };

    transaction.onerror = function () {

      reject(transaction.error);
    };

  });
}


// ---------------------------------
// WORD TO PDF TAB
// ---------------------------------

wordToPdfButton.addEventListener("click", () => {

  conversionMode = "word-to-pdf";

  wordToPdfButton.classList.add("active");
  pdfToWordButton.classList.remove("active");

  uploadTitle.textContent =
    "Upload your Word document";

  uploadDescription.textContent =
    "Drag & drop your DOC or DOCX file here";

  fileInput.accept =
    ".doc,.docx";

  fileInput.value = "";

  statusMessage.textContent = "";

});


// ---------------------------------
// PDF TO WORD TAB
// ---------------------------------

pdfToWordButton.addEventListener("click", () => {

  conversionMode = "pdf-to-word";

  pdfToWordButton.classList.add("active");
  wordToPdfButton.classList.remove("active");

  uploadTitle.textContent =
    "Upload your PDF document";

  uploadDescription.textContent =
    "Drag & drop your PDF file here";

  fileInput.accept = ".pdf";

  fileInput.value = "";

  statusMessage.textContent = "";

});


// ---------------------------------
// FILE SELECTED
// ---------------------------------

fileInput.addEventListener("change", () => {

  if (
    !fileInput.files ||
    fileInput.files.length === 0
  ) {
    return;
  }

  const file =
    fileInput.files[0];

  uploadTitle.textContent =
    file.name;

  uploadDescription.textContent =
    "File selected and ready to convert";

  statusMessage.textContent = "";

});


// ---------------------------------
// DRAG AND DROP
// ---------------------------------

uploadArea.addEventListener(
  "dragover",
  (event) => {

    event.preventDefault();

    uploadArea.classList.add(
      "dragging"
    );

  }
);


uploadArea.addEventListener(
  "dragleave",
  () => {

    uploadArea.classList.remove(
      "dragging"
    );

  }
);


uploadArea.addEventListener(
  "drop",
  (event) => {

    event.preventDefault();

    uploadArea.classList.remove(
      "dragging"
    );

    const file =
      event.dataTransfer.files[0];

    if (!file) {
      return;
    }

    const fileName =
      file.name.toLowerCase();


    if (
      conversionMode === "word-to-pdf" &&
      !fileName.endsWith(".doc") &&
      !fileName.endsWith(".docx")
    ) {

      statusMessage.textContent =
        "Please upload a DOC or DOCX file.";

      return;
    }


    if (
      conversionMode === "pdf-to-word" &&
      !fileName.endsWith(".pdf")
    ) {

      statusMessage.textContent =
        "Please upload a PDF file.";

      return;
    }


    const dataTransfer =
      new DataTransfer();

    dataTransfer.items.add(file);

    fileInput.files =
      dataTransfer.files;


    uploadTitle.textContent =
      file.name;

    uploadDescription.textContent =
      "File selected and ready to convert";

    statusMessage.textContent = "";

  }
);


// ---------------------------------
// CONVERT
// ---------------------------------

convertButton.addEventListener(
  "click",
  async () => {

    if (
      !fileInput.files ||
      fileInput.files.length === 0
    ) {

      statusMessage.textContent =
        "Please choose a file first.";

      return;
    }


    const file =
      fileInput.files[0];


    const maxSize =
      25 * 1024 * 1024;


    if (file.size > maxSize) {

      statusMessage.textContent =
        "File is too large. Maximum size is 25 MB.";

      return;
    }


    const formData =
      new FormData();

    formData.append(
      "file",
      file
    );


    const endpoint =
      conversionMode === "word-to-pdf"
        ? "/convert/word-to-pdf"
        : "/convert/pdf-to-word";


    convertButton.disabled = true;

    convertButton.textContent =
      "Converting...";

    statusMessage.textContent =
      "Converting your document...";


    try {

      const response =
        await fetch(
          endpoint,
          {
            method: "POST",
            body: formData
          }
        );


      if (!response.ok) {

        let message =
          "Conversion failed.";

        try {

          const errorData =
            await response.json();

          message =
            errorData.error || message;

        } catch (error) {

          console.error(error);
        }


        throw new Error(message);
      }


      const blob =
        await response.blob();


      // ---------------------------
      // WORD TO PDF
      // OPEN EDITOR
      // ---------------------------

      if (
        conversionMode === "word-to-pdf"
      ) {

        const originalName =
          file.name.replace(
            /\.(doc|docx)$/i,
            ""
          );

        const pdfName =
          originalName + ".pdf";


        statusMessage.textContent =
          "Opening PDF editor...";


        await savePdfForEditor(
          blob,
          pdfName
        );


        window.location.href =
          "/editor";

        return;
      }


      // ---------------------------
      // PDF TO WORD
      // KEEP NORMAL DOWNLOAD
      // ---------------------------

      const downloadURL =
        URL.createObjectURL(blob);


      const downloadLink =
        document.createElement("a");


      const originalName =
        file.name.replace(
          /\.pdf$/i,
          ""
        );


      downloadLink.href =
        downloadURL;


      downloadLink.download =
        originalName + ".docx";


      document.body.appendChild(
        downloadLink
      );


      downloadLink.click();

      downloadLink.remove();


      setTimeout(() => {

        URL.revokeObjectURL(
          downloadURL
        );

      }, 1000);


      statusMessage.textContent =
        "Conversion complete!";

    }

    catch (error) {

      console.error(error);

      statusMessage.textContent =
        error.message ||
        "Something went wrong during conversion.";

    }

    finally {

      convertButton.disabled = false;

      convertButton.textContent =
        "Convert File";

    }

  }
);
