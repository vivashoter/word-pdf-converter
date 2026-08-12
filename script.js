const wordToPdfButton = document.getElementById("wordToPdf");
const pdfToWordButton = document.getElementById("pdfToWord");

const fileInput = document.getElementById("fileInput");
const uploadTitle = document.getElementById("uploadTitle");
const uploadDescription = document.getElementById("uploadDescription");
const uploadArea = document.getElementById("uploadArea");

let conversionMode = "word-to-pdf";

wordToPdfButton.addEventListener("click", () => {
  conversionMode = "word-to-pdf";

  wordToPdfButton.classList.add("active");
  pdfToWordButton.classList.remove("active");

  uploadTitle.textContent = "Upload your Word document";
  uploadDescription.textContent = "Drag & drop your DOC or DOCX file here";

  fileInput.accept = ".doc,.docx";
  fileInput.value = "";
});

pdfToWordButton.addEventListener("click", () => {
  conversionMode = "pdf-to-word";

  pdfToWordButton.classList.add("active");
  wordToPdfButton.classList.remove("active");

  uploadTitle.textContent = "Upload your PDF document";
  uploadDescription.textContent = "Drag & drop your PDF file here";

  fileInput.accept = ".pdf";
  fileInput.value = "";
});

fileInput.addEventListener("change", () => {
  if (fileInput.files.length > 0) {
    const file = fileInput.files[0];

    uploadTitle.textContent = file.name;
    uploadDescription.textContent = "File selected and ready to convert";
  }
});

uploadArea.addEventListener("dragover", (event) => {
  event.preventDefault();
  uploadArea.classList.add("dragging");
});

uploadArea.addEventListener("dragleave", () => {
  uploadArea.classList.remove("dragging");
});

uploadArea.addEventListener("drop", (event) => {
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
    alert("Please upload a DOC or DOCX file.");
    return;
  }

  if (
    conversionMode === "pdf-to-word" &&
    !fileName.endsWith(".pdf")
  ) {
    alert("Please upload a PDF file.");
    return;
  }

  const dataTransfer = new DataTransfer();
  dataTransfer.items.add(file);
  fileInput.files = dataTransfer.files;

  uploadTitle.textContent = file.name;
  uploadDescription.textContent = "File selected and ready to convert";
});
