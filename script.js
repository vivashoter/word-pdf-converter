const wordToPdfButton =
  document.getElementById("wordToPdf");

const pdfToWordButton =
  document.getElementById("pdfToWord");

const fileInput =
  document.getElementById("fileInput");

const uploadTitle =
  document.getElementById("uploadTitle");

const uploadDescription =
  document.getElementById("uploadDescription");

const uploadArea =
  document.getElementById("uploadArea");

const convertButton =
  document.getElementById("convertButton");

const statusMessage =
  document.getElementById("statusMessage");


let conversionMode =
  "word-to-pdf";


// --------------------------------
// STATUS
// --------------------------------

function setStatus(
  message,
  type = ""
) {

  statusMessage.textContent =
    message;


  statusMessage.classList.remove(
    "success",
    "error"
  );


  if (type) {

    statusMessage.classList.add(
      type
    );

  }

}


// --------------------------------
// RESET FILE
// --------------------------------

function resetFileSelection() {

  fileInput.value = "";

  setStatus("");

}


// --------------------------------
// WORD TO PDF
// --------------------------------

wordToPdfButton.addEventListener(
  "click",
  () => {

    conversionMode =
      "word-to-pdf";


    wordToPdfButton.classList.add(
      "active"
    );


    pdfToWordButton.classList.remove(
      "active"
    );


    uploadTitle.textContent =
      "Upload your Word document";


    uploadDescription.textContent =
      "Drag & drop your DOC or DOCX file here";


    fileInput.accept =
      ".doc,.docx";


    resetFileSelection();

  }
);


// --------------------------------
// PDF TO WORD
// --------------------------------

pdfToWordButton.addEventListener(
  "click",
  () => {

    conversionMode =
      "pdf-to-word";


    pdfToWordButton.classList.add(
      "active"
    );


    wordToPdfButton.classList.remove(
      "active"
    );


    uploadTitle.textContent =
      "Upload your PDF document";


    uploadDescription.textContent =
      "Drag & drop your PDF file here";


    fileInput.accept =
      ".pdf";


    resetFileSelection();

  }
);


// --------------------------------
// VALIDATE FILE
// --------------------------------

function validateFile(file) {

  if (!file) {

    return {
      valid: false,
      message:
        "Please choose a file first."
    };

  }


  const fileName =
    file.name.toLowerCase();


  if (
    conversionMode ===
    "word-to-pdf"
  ) {

    if (
      !fileName.endsWith(".doc") &&
      !fileName.endsWith(".docx")
    ) {

      return {
        valid: false,
        message:
          "Please choose a DOC or DOCX file."
      };

    }

  }


  if (
    conversionMode ===
    "pdf-to-word"
  ) {

    if (
      !fileName.endsWith(".pdf")
    ) {

      return {
        valid: false,
        message:
          "Please choose a PDF file."
      };

    }

  }


  const maxSize =
    25 * 1024 * 1024;


  if (
    file.size > maxSize
  ) {

    return {
      valid: false,
      message:
        "File is too large. Maximum size is 25 MB."
    };

  }


  return {
    valid: true
  };

}


// --------------------------------
// SHOW SELECTED FILE
// --------------------------------

function showSelectedFile(file) {

  const validation =
    validateFile(file);


  if (!validation.valid) {

    setStatus(
      validation.message,
      "error"
    );

    return false;
  }


  uploadTitle.textContent =
    file.name;


  uploadDescription.textContent =
    "File selected and ready to convert";


  setStatus("");


  return true;

}


// --------------------------------
// FILE PICKER
// --------------------------------

fileInput.addEventListener(
  "change",
  () => {

    if (
      !fileInput.files ||
      fileInput.files.length === 0
    ) {

      return;
    }


    showSelectedFile(
      fileInput.files[0]
    );

  }
);


// --------------------------------
// KEYBOARD ACCESS
// --------------------------------

uploadArea.addEventListener(
  "keydown",
  (event) => {

    if (
      event.key === "Enter" ||
      event.key === " "
    ) {

      const target =
        event.target;


      if (
        target === uploadArea
      ) {

        event.preventDefault();

        fileInput.click();

      }

    }

  }
);


// --------------------------------
// DRAG OVER
// --------------------------------

uploadArea.addEventListener(
  "dragover",
  (event) => {

    event.preventDefault();


    uploadArea.classList.add(
      "dragging"
    );

  }
);


// --------------------------------
// DRAG LEAVE
// --------------------------------

uploadArea.addEventListener(
  "dragleave",
  () => {

    uploadArea.classList.remove(
      "dragging"
    );

  }
);


// --------------------------------
// DROP
// --------------------------------

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


    if (
      !showSelectedFile(file)
    ) {

      return;

    }


    try {

      const transfer =
        new DataTransfer();


      transfer.items.add(file);


      fileInput.files =
        transfer.files;

    }

    catch (error) {

      console.warn(
        "Drag-and-drop file assignment is not supported in this browser.",
        error
      );

    }

  }
);


// --------------------------------
// DOWNLOAD RESPONSE
// --------------------------------

function downloadConvertedFile(
  blob,
  outputName
) {

  const downloadURL =
    URL.createObjectURL(blob);


  const downloadLink =
    document.createElement("a");


  downloadLink.href =
    downloadURL;


  downloadLink.download =
    outputName;


  downloadLink.style.display =
    "none";


  document.body.appendChild(
    downloadLink
  );


  downloadLink.click();


  downloadLink.remove();


  setTimeout(
    () => {

      URL.revokeObjectURL(
        downloadURL
      );

    },
    3000
  );

}


// --------------------------------
// CONVERT
// --------------------------------

convertButton.addEventListener(
  "click",
  async () => {

    const file =
      fileInput.files?.[0];


    const validation =
      validateFile(file);


    if (!validation.valid) {

      setStatus(
        validation.message,
        "error"
      );

      return;

    }


    const formData =
      new FormData();


    formData.append(
      "file",
      file
    );


    const endpoint =
      conversionMode ===
      "word-to-pdf"
        ? "/convert/word-to-pdf"
        : "/convert/pdf-to-word";


    convertButton.disabled =
      true;


    convertButton.textContent =
      "Converting...";


    setStatus(
      "Converting your document..."
    );


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


          if (
            errorData.error
          ) {

            message =
              errorData.error;

          }

        }

        catch (error) {

          console.error(error);

        }


        throw new Error(
          message
        );

      }


      const blob =
        await response.blob();


      let outputName;


      if (
        conversionMode ===
        "word-to-pdf"
      ) {

        const originalName =
          file.name.replace(
            /\.(doc|docx)$/i,
            ""
          );


        outputName =
          originalName +
          ".pdf";

      }

      else {

        const originalName =
          file.name.replace(
            /\.pdf$/i,
            ""
          );


        outputName =
          originalName +
          ".docx";

      }


      downloadConvertedFile(
        blob,
        outputName
      );


      setStatus(
        "Conversion complete! Your file has been downloaded.",
        "success"
      );

    }

    catch (error) {

      console.error(error);


      setStatus(
        error.message ||
        "Something went wrong during conversion.",
        "error"
      );

    }

    finally {

      convertButton.disabled =
        false;


      convertButton.textContent =
        "Convert File";

    }

  }
);
