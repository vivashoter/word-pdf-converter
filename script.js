const wordToPdfButton =
    document.getElementById("wordToPdf");

const pdfToWordButton =
    document.getElementById("pdfToWord");

const converterTabs =
    document.getElementById("converterTabs");

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

const progressContainer =
    document.getElementById("progressContainer");

const successArea =
    document.getElementById("successArea");

const convertedFileName =
    document.getElementById("convertedFileName");

const downloadAgainButton =
    document.getElementById("downloadAgainButton");

const convertAnotherButton =
    document.getElementById("convertAnotherButton");


let conversionMode =
    "word-to-pdf";


let convertedBlob = null;

let convertedOutputName = "";


// --------------------------------
// STATUS
// --------------------------------

function setStatus(message, type = "") {

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
// RESET UPLOAD TEXT
// --------------------------------

function resetUploadText() {

    if (
        conversionMode ===
        "word-to-pdf"
    ) {

        uploadTitle.textContent =
            "Upload your Word document";


        uploadDescription.textContent =
            "Drag & drop your DOC or DOCX file here";

    } else {

        uploadTitle.textContent =
            "Upload your PDF document";


        uploadDescription.textContent =
            "Drag & drop your PDF file here";

    }

}


// --------------------------------
// RESET FILE
// --------------------------------

function resetFileSelection() {

    fileInput.value = "";


    progressContainer.hidden =
        true;


    setStatus("");


    resetUploadText();

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


        fileInput.accept =
            ".pdf";


        resetFileSelection();

    }
);


// --------------------------------
// VALIDATE
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


    const maximumFileSize =
        25 * 1024 * 1024;


    if (
        file.size >
        maximumFileSize
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
// SELECTED FILE
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


    progressContainer.hidden =
        true;


    setStatus("");


    return true;

}


// --------------------------------
// FILE INPUT
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
// KEYBOARD
// --------------------------------

uploadArea.addEventListener(
    "keydown",
    (event) => {

        if (
            event.target === uploadArea &&
            (
                event.key === "Enter" ||
                event.key === " "
            )
        ) {

            event.preventDefault();


            fileInput.click();

        }

    }
);


// --------------------------------
// DRAG
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


            transfer.items.add(
                file
            );


            fileInput.files =
                transfer.files;

        } catch (error) {

            console.warn(
                "Drag-and-drop file assignment is not supported in this browser.",
                error
            );

        }

    }
);


// --------------------------------
// DOWNLOAD
// --------------------------------

function downloadConvertedFile(
    blob,
    outputName
) {

    const downloadUrl =
        URL.createObjectURL(blob);


    const link =
        document.createElement("a");


    link.href =
        downloadUrl;


    link.download =
        outputName;


    link.style.display =
        "none";


    document.body.appendChild(
        link
    );


    link.click();


    link.remove();


    setTimeout(
        () => {

            URL.revokeObjectURL(
                downloadUrl
            );

        },
        5000
    );

}


// --------------------------------
// SHOW SUCCESS
// --------------------------------

function showSuccessScreen(
    blob,
    outputName
) {

    convertedBlob =
        blob;


    convertedOutputName =
        outputName;


    convertedFileName.textContent =
        outputName;


    converterTabs.hidden =
        true;


    uploadArea.hidden =
        true;


    successArea.hidden =
        false;

}


// --------------------------------
// DOWNLOAD AGAIN
// --------------------------------

downloadAgainButton.addEventListener(
    "click",
    () => {

        if (
            !convertedBlob ||
            !convertedOutputName
        ) {

            return;

        }


        downloadConvertedFile(
            convertedBlob,
            convertedOutputName
        );

    }
);


// --------------------------------
// CONVERT ANOTHER
// --------------------------------

convertAnotherButton.addEventListener(
    "click",
    () => {

        convertedBlob =
            null;


        convertedOutputName =
            "";


        successArea.hidden =
            true;


        converterTabs.hidden =
            false;


        uploadArea.hidden =
            false;


        resetFileSelection();

    }
);


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


        wordToPdfButton.disabled =
            true;


        pdfToWordButton.disabled =
            true;


        fileInput.disabled =
            true;


        convertButton.textContent =
            "Converting...";


        progressContainer.hidden =
            false;


        setStatus("");


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

                let errorMessage =
                    "Conversion failed. Please try again.";


                try {

                    const errorData =
                        await response.json();


                    if (
                        errorData &&
                        errorData.error
                    ) {

                        errorMessage =
                            errorData.error;

                    }

                } catch (error) {

                    console.error(
                        "Could not read server error:",
                        error
                    );

                }


                throw new Error(
                    errorMessage
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

            } else {

                const originalName =
                    file.name.replace(
                        /\.pdf$/i,
                        ""
                    );


                outputName =
                    originalName +
                    ".docx";

            }


            progressContainer.hidden =
                true;


            downloadConvertedFile(
                blob,
                outputName
            );


            showSuccessScreen(
                blob,
                outputName
            );

        } catch (error) {

            console.error(
                "Conversion error:",
                error
            );


            setStatus(
                error.message ||
                "Something went wrong during conversion.",
                "error"
            );

        } finally {

            progressContainer.hidden =
                true;


            convertButton.disabled =
                false;


            wordToPdfButton.disabled =
                false;


            pdfToWordButton.disabled =
                false;


            fileInput.disabled =
                false;


            convertButton.textContent =
                "Convert File";

        }

    }
);
