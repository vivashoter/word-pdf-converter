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

const progressBar =
    document.getElementById("progressBar");

const progressTrack =
    document.getElementById("progressTrack");

const progressLabel =
    document.getElementById("progressLabel");

const progressStatus =
    document.getElementById("progressStatus");

const progressNote =
    document.getElementById("progressNote");

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

let progressTimer = null;

let progressValue = 0;


// ========================================
// STATUS
// ========================================

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


// ========================================
// MODE APPEARANCE
// ========================================

function updateModeAppearance() {

    uploadArea.classList.remove(
        "word-mode",
        "pdf-mode"
    );


    progressContainer.classList.remove(
        "pdf-progress"
    );


    if (
        conversionMode ===
        "word-to-pdf"
    ) {

        uploadArea.classList.add(
            "word-mode"
        );

        progressLabel.textContent =
            "Converting Word to PDF…";

    }

    else {

        uploadArea.classList.add(
            "pdf-mode"
        );

        progressContainer.classList.add(
            "pdf-progress"
        );

        progressLabel.textContent =
            "Converting PDF to Word…";

    }

}


// ========================================
// RESET UPLOAD TEXT
// ========================================

function resetUploadText() {

    if (
        conversionMode ===
        "word-to-pdf"
    ) {

        uploadTitle.textContent =
            "Upload your Word document";


        uploadDescription.textContent =
            "Drag & drop your DOC or DOCX file here";

    }

    else {

        uploadTitle.textContent =
            "Upload your PDF document";


        uploadDescription.textContent =
            "Drag & drop your PDF file here";

    }

}


// ========================================
// RESET PROGRESS
// ========================================

function resetProgress() {

    if (progressTimer) {

        clearInterval(
            progressTimer
        );

        progressTimer = null;

    }


    progressValue = 0;


    progressBar.style.width =
        "0%";


    progressTrack.setAttribute(
        "aria-valuenow",
        "0"
    );


    progressStatus.textContent =
        "Processing";


    progressNote.textContent =
        "Please keep this page open while your document is converted.";

}


// ========================================
// START CONVERSION PROGRESS
// ========================================

function startProgress() {

    resetProgress();


    progressContainer.hidden =
        false;


    progressValue = 6;


    progressBar.style.width =
        progressValue + "%";


    const startTime =
        Date.now();


    progressTimer =
        setInterval(
            () => {

                const elapsedSeconds =
                    (
                        Date.now() -
                        startTime
                    ) / 1000;


                /*
                   The server does not provide an exact
                   conversion percentage.

                   This indicator advances with the real
                   elapsed processing time but intentionally
                   stops below 100 until the server confirms
                   completion.
                */


                let target;


                if (
                    elapsedSeconds < 2
                ) {

                    target =
                        18;

                    progressStatus.textContent =
                        "Uploading";

                }

                else if (
                    elapsedSeconds < 5
                ) {

                    target =
                        38;

                    progressStatus.textContent =
                        "Reading document";

                }

                else if (
                    elapsedSeconds < 10
                ) {

                    target =
                        58;

                    progressStatus.textContent =
                        "Converting";

                }

                else if (
                    elapsedSeconds < 20
                ) {

                    target =
                        74;

                    progressStatus.textContent =
                        "Converting";

                }

                else if (
                    elapsedSeconds < 35
                ) {

                    target =
                        84;

                    progressStatus.textContent =
                        "Finishing up";

                }

                else {

                    target =
                        92;

                    progressStatus.textContent =
                        "Still working";

                }


                /*
                   Smoothly approach the target rather
                   than jumping immediately.
                */

                if (
                    progressValue <
                    target
                ) {

                    progressValue +=
                        Math.max(
                            0.7,
                            (
                                target -
                                progressValue
                            ) * 0.10
                        );


                    if (
                        progressValue >
                        target
                    ) {

                        progressValue =
                            target;

                    }


                    progressBar.style.width =
                        progressValue +
                        "%";


                    progressTrack.setAttribute(
                        "aria-valuenow",
                        Math.round(
                            progressValue
                        )
                    );

                }

            },
            500
        );

}


// ========================================
// COMPLETE PROGRESS
// ========================================

function completeProgress() {

    if (progressTimer) {

        clearInterval(
            progressTimer
        );

        progressTimer =
            null;

    }


    progressValue =
        100;


    progressStatus.textContent =
        "Complete";


    progressBar.style.width =
        "100%";


    progressTrack.setAttribute(
        "aria-valuenow",
        "100"
    );


    progressNote.textContent =
        "Your converted file is ready.";

}


// ========================================
// RESET FILE
// ========================================

function resetFileSelection() {

    fileInput.value = "";


    progressContainer.hidden =
        true;


    resetProgress();


    setStatus("");


    resetUploadText();


    updateModeAppearance();

}


// ========================================
// WORD TO PDF
// ========================================

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


// ========================================
// PDF TO WORD
// ========================================

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


// ========================================
// VALIDATE
// ========================================

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


// ========================================
// SELECT FILE
// ========================================

function showSelectedFile(file) {

    const validation =
        validateFile(file);


    if (
        !validation.valid
    ) {

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


// ========================================
// FILE INPUT
// ========================================

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


// ========================================
// DRAG AND DROP
// ========================================

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

        }

        catch (error) {

            console.warn(
                "Dropped file could not be assigned.",
                error
            );

        }

    }
);


// ========================================
// DOWNLOAD
// ========================================

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


// ========================================
// SUCCESS
// ========================================

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


// ========================================
// DOWNLOAD AGAIN
// ========================================

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


// ========================================
// CONVERT ANOTHER
// ========================================

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


// ========================================
// CONVERT
// ========================================

convertButton.addEventListener(
    "click",
    async () => {

        const file =
            fileInput.files?.[0];


        const validation =
            validateFile(file);


        if (
            !validation.valid
        ) {

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


        updateModeAppearance();


        startProgress();


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


            if (
                !response.ok
            ) {

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

                }

                catch (error) {

                    console.error(
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

                outputName =
                    file.name.replace(
                        /\.(doc|docx)$/i,
                        ""
                    ) +
                    ".pdf";

            }

            else {

                outputName =
                    file.name.replace(
                        /\.pdf$/i,
                        ""
                    ) +
                    ".docx";

            }


            completeProgress();


            /*
               Allow the user to actually see the
               progress bar reach 100 before the
               success screen replaces it.
            */

            await new Promise(
                resolve =>
                    setTimeout(
                        resolve,
                        650
                    )
            );


            downloadConvertedFile(
                blob,
                outputName
            );


            showSuccessScreen(
                blob,
                outputName
            );

        }

        catch (error) {

            resetProgress();


            progressContainer.hidden =
                true;


            setStatus(
                error.message ||
                "Something went wrong during conversion.",
                "error"
            );

        }

        finally {

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


// ========================================
// INITIAL STATE
// ========================================

updateModeAppearance();
