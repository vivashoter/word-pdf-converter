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
}

const convertButton = document.getElementById("convertButton");
const statusMessage = document.getElementById("statusMessage");

convertButton.addEventListener("click", async () => {
  if (fileInput.files.length === 0) {
    statusMessage.textContent = "Please choose a file first.";
    return;
  }

  const file = fileInput.files[0];

  // 25 MB limit
  const maxSize = 25 * 1024 * 1024;

  if (file.size > maxSize) {
    statusMessage.textContent = "File is too large. Maximum size is 25 MB.";
    return;
  }

  // PDF to Word will be added next
  if (conversionMode === "pdf-to-word") {
    statusMessage.textContent = "PDF to Word conversion is coming next.";
    return;
  }

  const formData = new FormData();
  formData.append("file", file);

  convertButton.disabled = true;
  convertButton.textContent = "Converting...";
  statusMessage.textContent = "Converting your document...";

  try {
    const response = await fetch("/convert/word-to-pdf", {
      method: "POST",
      body: formData
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error || "Conversion failed.");
    }

    const blob = await response.blob();

    const downloadURL = URL.createObjectURL(blob);

    const downloadLink = document.createElement("a");

    downloadLink.href = downloadURL;

    const originalName = file.name.replace(/\.(doc|docx)$/i, "");
    downloadLink.download = originalName + ".pdf";

    document.body.appendChild(downloadLink);

    downloadLink.click();

    downloadLink.remove();

    URL.revokeObjectURL(downloadURL);

    statusMessage.textContent = "Conversion complete! Your PDF has been downloaded.";

  } catch (error) {
    console.error(error);

    statusMessage.textContent =
      error.message || "Something went wrong during conversion.";

  } finally {
    convertButton.disabled = false;
    convertButton.textContent = "Convert File";
  }
});
