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
    !file
