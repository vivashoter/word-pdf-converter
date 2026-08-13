let zoomLevel = 100;

const pdfViewer = document.getElementById("pdfViewer");
const zoomLevelText = document.getElementById("zoomLevel");

const zoomInButton = document.getElementById("zoomIn");
const zoomOutButton = document.getElementById("zoomOut");

const printButton = document.getElementById("printButton");
const downloadButton = document.getElementById("downloadButton");
const emailButton = document.getElementById("emailButton");

const mobilePrint = document.getElementById("mobilePrint");
const mobileDownload = document.getElementById("mobileDownload");
const mobileEmail = document.getElementById("mobileEmail");

const documentName = document.getElementById("documentName");


// ZOOM
function updateZoom() {
  pdfViewer.style.transform = `scale(${zoomLevel / 100})`;
  zoomLevelText.textContent = `${zoomLevel}%`;
}

zoomInButton.addEventListener("click", () => {
  if (zoomLevel < 200) {
    zoomLevel += 10;
    updateZoom();
  }
});

zoomOutButton.addEventListener("click", () => {
  if (zoomLevel > 50) {
    zoomLevel -= 10;
    updateZoom();
  }
});


// PRINT
function printDocument() {
  window.print();
}

printButton.addEventListener("click", printDocument);
mobilePrint.addEventListener("click", printDocument);


// DOWNLOAD
function downloadDocument() {
  alert("The converted PDF will be connected to this button next.");
}

downloadButton.addEventListener("click", downloadDocument);
mobileDownload.addEventListener("click", downloadDocument);


// EMAIL
function emailDocument() {
  alert("Email sharing will be connected after the PDF editor is working.");
}

emailButton.addEventListener("click", emailDocument);
mobileEmail.addEventListener("click", emailDocument);


// TOOLBAR
const toolButtons = document.querySelectorAll(".tool-button");

toolButtons.forEach((button) => {
  button.addEventListener("click", () => {
    toolButtons.forEach((item) => {
      item.classList.remove("active-tool");
    });

    button.classList.add("active-tool");
  });
});


// DEFAULT ZOOM
updateZoom();
