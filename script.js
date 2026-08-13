const wordToPdfButton = document.getElementById("wordToPdf");
const pdfToWordButton = document.getElementById("pdfToWord");

const fileInput = document.getElementById("fileInput");
const uploadTitle = document.getElementById("uploadTitle");
const uploadDescription = document.getElementById("uploadDescription");
const uploadArea = document.getElementById("uploadArea");

const convertButton = document.getElementById("convertButton");
const statusMessage = document.getElementById("statusMessage");

let conversionMode = "word-to-pdf";


function resetFileSelection() {
  fileInput.value = "";
  statusMessage.textContent = "";
}


wordToPdfButton.addEventListener("click", function () {
  conversionMode = "word-to-pdf";

  wordToPdfButton.classList.add("active");
  pdfToWordButton.classList.remove("active");

  uploadTitle.textContent = "Upload your Word document";
  uploadDescription.textContent =
    "Drag & drop your DOC or DOCX file here";

  fileInput.accept = ".doc,.docx";

  resetFileSelection();
});


pdfToWordButton.addEventListener("click", function () {
  conversionMode = "pdf-to-word";

  pdfToWordButton.classList.add("active");
  wordToPdfButton.classList.remove("active");

  uploadTitle.textContent = "Upload your PDF document";
  uploadDescription.textContent =
    "Drag & drop your PDF file here";

  fileInput.accept = ".pdf";

  resetFileSelection();
});


fileInput.addEventListener("change", function () {
  if (!this.files || this.files.length === 0) {
    return;
  }

  const file = this.files[0];

  uploadTitle.textContent = file.name;
  uploadDescription.textContent = "File selected and ready to convert";

  statusMessage.textContent = "";
});


uploadArea.addEventListener("dragover", function (event) {
  event.preventDefault();
  uploadArea.classList.add("dragging");
});


uploadArea.addEventListener("dragleave", function () {
  uploadArea.classList.remove("dragging");
});


uploadArea.addEventListener("drop", function (event) {
  event.preventDefault();

  uploadArea.classList.remove("dragging");

  const file = event.dataTransfer.files[0];

  if (!file) {
    return;
  }

  const fileName = file.name.toLowerCase();

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

  const dataTransfer = new DataTransfer();

  dataTransfer.items.add(file);

  fileInput.files = dataTransfer.files;

  uploadTitle.textContent = file.name;
  uploadDescription.textContent =
    "File selected and ready to convert";
});


convertButton.addEventListener("click", async function () {
  if (!fileInput.files || fileInput.files.length === 0) {
    statusMessage.textContent =
      "Please choose a file first.";
    return;
  }

  const file = fileInput.files[0];

  const maxSize = 25 * 1024 * 1024;

  if (file.size > maxSize) {
    statusMessage.textContent =
      "File is too large. Maximum size is 25 MB.";
    return;
  }

  if (conversionMode === "pdf-to-word") {
    statusMessage.textContent =
      "PDF to Word conversion is not connected yet.";
    return;
  }

  const formData = new FormData();

  formData.append("file", file);

  convertButton.disabled = true;
  convertButton.textContent = "Converting...";

  statusMessage.textContent =
    "Converting your document...";

  try {
    const response = await fetch("/convert/word-to-pdf", {
      method: "POST",
      body: formData
    });

    if (!response.ok) {
      let message = "Conversion failed.";

      try {
        const errorData = await response.json();

        if (errorData.error) {
          message = errorData.error;
        }
      } catch (error) {
        console.error(error);
      }

      throw new Error(message);
    }

    const blob = await response.blob();

    const downloadURL = URL.createObjectURL(blob);

    const downloadLink = document.createElement("a");

    downloadLink.href = downloadURL;

    const originalName =
      file.name.replace(/\.(doc|docx)$/i, "");

    downloadLink.download =
      originalName + ".pdf";

    document.body.appendChild(downloadLink);

    downloadLink.click();

    downloadLink.remove();

    URL.revokeObjectURL(downloadURL);

    statusMessage.textContent =
      "Conversion complete!";

  } catch (error) {
    console.error(error);

    statusMessage.textContent =
      error.message || "Something went wrong.";

  } finally {
    convertButton.disabled = false;
    convertButton.textContent = "Convert File";
  }
});
