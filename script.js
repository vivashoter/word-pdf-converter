/* ========================================
   GOOGLE CONFIG
======================================== */

const GOOGLE_CLIENT_ID =
    "PASTE_YOUR_CLIENT_ID_HERE";

const GOOGLE_API_KEY =
    "PASTE_YOUR_API_KEY_HERE";

const GOOGLE_APP_ID =
    "239816509439";

const GOOGLE_SCOPE =
    "https://www.googleapis.com/auth/drive.file";


/* ========================================
   MIME TYPES
======================================== */

const GOOGLE_DOC_MIME =
    "application/vnd.google-apps.document";

const DOCX_MIME =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const DOC_MIME =
    "application/msword";

const PDF_MIME =
    "application/pdf";


/* ========================================
   ELEMENTS
======================================== */

const wordToPdfButton =
    document.getElementById("wordToPdf");

const pdfToWordButton =
    document.getElementById("pdfToWord");

const googleToPdfButton =
    document.getElementById("googleToPdf");

const googleToWordButton =
    document.getElementById("googleToWord");

const converterTabs =
    document.getElementById("converterTabs");

const fileInput =
    document.getElementById("fileInput");

const uploadArea =
    document.getElementById("uploadArea");

const googleArea =
    document.getElementById("googleArea");

const uploadTitle =
    document.getElementById("uploadTitle");

const uploadDescription =
    document.getElementById("uploadDescription");

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


/* ========================================
   GOOGLE ELEMENTS
======================================== */

const googleTitle =
    document.getElementById("googleTitle");

const googleDescription =
    document.getElementById("googleDescription");

const chooseGoogleDocButton =
    document.getElementById("chooseGoogleDocButton");

const uploadGoogleDocButton =
    document.getElementById("uploadGoogleDocButton");

const selectedGoogleFile =
    document.getElementById("selectedGoogleFile");

const selectedGoogleFileName =
    document.getElementById("selectedGoogleFileName");

const convertGoogleButton =
    document.getElementById("convertGoogleButton");

const googleStatusMessage =
    document.getElementById("googleStatusMessage");

const googleProgressContainer =
    document.getElementById("googleProgressContainer");

const googleProgressBar =
    document.getElementById("googleProgressBar");

const googleProgressTrack =
    document.getElementById("googleProgressTrack");

const googleProgressLabel =
    document.getElementById("googleProgressLabel");

const googleProgressStatus =
    document.getElementById("googleProgressStatus");


/* ========================================
   GENERAL STATE
======================================== */

let conversionMode =
    "word-to-pdf";

let convertedBlob =
    null;

let convertedOutputName =
    "";

let progressTimer =
    null;

let progressValue =
    0;


/* ========================================
   GOOGLE STATE
======================================== */

let googleTokenClient =
    null;

let googleAccessToken =
    null;

let googlePickerReady =
    false;

let googleIdentityReady =
    false;

let requestedPickerType =
    "drive";

let selectedGoogleDocId =
    null;

let selectedGoogleDocName =
    null;

let selectedGoogleDocMimeType =
    null;

let googleProgressTimer =
    null;

let googleProgressValue =
    0;


/* ========================================
   STATUS
======================================== */

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


function setGoogleStatus(
    message,
    type = ""
) {

    googleStatusMessage.textContent =
        message;

    googleStatusMessage.classList.remove(
        "success",
        "error"
    );

    if (type) {

        googleStatusMessage.classList.add(
            type
        );

    }

}


/* ========================================
   TAB HELPERS
======================================== */

function clearActiveTabs() {

    wordToPdfButton.classList.remove(
        "active"
    );

    pdfToWordButton.classList.remove(
        "active"
    );

    googleToPdfButton.classList.remove(
        "active"
    );

    googleToWordButton.classList.remove(
        "active"
    );

}


function disableAllTabs(
    disabled
) {

    wordToPdfButton.disabled =
        disabled;

    pdfToWordButton.disabled =
        disabled;

    googleToPdfButton.disabled =
        disabled;

    googleToWordButton.disabled =
        disabled;

}


/* ========================================
   SET MODE
======================================== */

function setMode(
    mode
) {

    conversionMode =
        mode;

    clearActiveTabs();

    successArea.hidden =
        true;

    converterTabs.hidden =
        false;


    if (
        mode === "word-to-pdf"
    ) {

        wordToPdfButton.classList.add(
            "active"
        );

        uploadArea.hidden =
            false;

        googleArea.hidden =
            true;

        uploadArea.className =
            "upload-area word-mode";

        fileInput.accept =
            ".doc,.docx";

        resetFileSelection();

    }


    else if (
        mode === "pdf-to-word"
    ) {

        pdfToWordButton.classList.add(
            "active"
        );

        uploadArea.hidden =
            false;

        googleArea.hidden =
            true;

        uploadArea.className =
            "upload-area pdf-mode";

        fileInput.accept =
            ".pdf";

        resetFileSelection();

    }


    else if (
        mode === "google-to-pdf"
    ) {

        googleToPdfButton.classList.add(
            "active"
        );

        uploadArea.hidden =
            true;

        googleArea.hidden =
            false;

        googleTitle.textContent =
            "Google Doc to PDF";

        googleDescription.textContent =
            "Choose a Google Doc from Drive or upload a Word document from your computer.";

        resetGoogleSelection();

    }


    else if (
        mode === "google-to-word"
    ) {

        googleToWordButton.classList.add(
            "active"
        );

        uploadArea.hidden =
            true;

        googleArea.hidden =
            false;

        googleTitle.textContent =
            "Google Doc to Word";

        googleDescription.textContent =
            "Choose a Google Doc from Drive or upload a PDF from your computer.";

        resetGoogleSelection();

    }

}


/* ========================================
   TABS
======================================== */

wordToPdfButton.addEventListener(
    "click",
    () => {

        setMode(
            "word-to-pdf"
        );

    }
);


pdfToWordButton.addEventListener(
    "click",
    () => {

        setMode(
            "pdf-to-word"
        );

    }
);


googleToPdfButton.addEventListener(
    "click",
    () => {

        setMode(
            "google-to-pdf"
        );

    }
);


googleToWordButton.addEventListener(
    "click",
    () => {

        setMode(
            "google-to-word"
        );

    }
);


/* ========================================
   LOCAL FILE RESET
======================================== */

function resetFileSelection() {

    fileInput.value =
        "";

    progressContainer.hidden =
        true;

    resetProgress();

    setStatus("");


    if (
        conversionMode ===
        "word-to-pdf"
    ) {

        uploadTitle.textContent =
            "Upload your Word document";

        uploadDescription.textContent =
            "Drag & drop your DOC or DOCX file here";

        progressLabel.textContent =
            "Converting Word to PDF…";

    }

    else {

        uploadTitle.textContent =
            "Upload your PDF document";

        uploadDescription.textContent =
            "Drag & drop your PDF file here";

        progressLabel.textContent =
            "Converting PDF to Word…";

    }

}


/* ========================================
   LOCAL VALIDATION
======================================== */

function validateFile(
    file
) {

    if (!file) {

        return {
            valid: false,
            message:
                "Please choose a file first."
        };

    }


    const name =
        file.name.toLowerCase();


    if (
        conversionMode ===
        "word-to-pdf"
    ) {

        if (
            !name.endsWith(".doc") &&
            !name.endsWith(".docx")
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
            !name.endsWith(".pdf")
        ) {

            return {
                valid: false,
                message:
                    "Please choose a PDF file."
            };

        }

    }


    if (
        file.size >
        25 * 1024 * 1024
    ) {

        return {
            valid: false,
            message:
                "Maximum file size is 25 MB."
        };

    }


    return {
        valid: true
    };

}


/* ========================================
   LOCAL SELECT
======================================== */

function showSelectedFile(
    file
) {

    const validation =
        validateFile(
            file
        );


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

    setStatus("");

    return true;

}


fileInput.addEventListener(
    "change",
    () => {

        const file =
            fileInput.files?.[0];

        if (file) {

            showSelectedFile(
                file
            );

        }

    }
);


/* ========================================
   DRAG DROP
======================================== */

uploadArea.addEventListener(
    "dragover",
    event => {

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
    event => {

        event.preventDefault();

        uploadArea.classList.remove(
            "dragging"
        );


        const file =
            event.dataTransfer.files?.[0];


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

            console.error(
                error
            );

        }

    }
);


/* ========================================
   NORMAL PROGRESS
======================================== */

function resetProgress() {

    if (
        progressTimer
    ) {

        clearInterval(
            progressTimer
        );

    }


    progressTimer =
        null;

    progressValue =
        0;

    progressBar.style.width =
        "0%";

    progressTrack.setAttribute(
        "aria-valuenow",
        "0"
    );

    progressStatus.textContent =
        "Processing";

}


function startProgress() {

    resetProgress();

    progressContainer.hidden =
        false;

    progressValue =
        6;

    progressBar.style.width =
        "6%";


    const started =
        Date.now();


    progressTimer =
        setInterval(
            () => {

                const seconds =
                    (
                        Date.now() -
                        started
                    ) / 1000;


                let target;


                if (
                    seconds < 2
                ) {

                    target =
                        18;

                    progressStatus.textContent =
                        "Uploading";

                }

                else if (
                    seconds < 5
                ) {

                    target =
                        38;

                    progressStatus.textContent =
                        "Reading";

                }

                else if (
                    seconds < 10
                ) {

                    target =
                        58;

                    progressStatus.textContent =
                        "Converting";

                }

                else if (
                    seconds < 20
                ) {

                    target =
                        74;

                    progressStatus.textContent =
                        "Converting";

                }

                else {

                    target =
                        92;

                    progressStatus.textContent =
                        "Finishing";

                }


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
                            ) *
                            0.10
                        );


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


function completeProgress() {

    if (
        progressTimer
    ) {

        clearInterval(
            progressTimer
        );

    }


    progressTimer =
        null;

    progressValue =
        100;

    progressBar.style.width =
        "100%";

    progressTrack.setAttribute(
        "aria-valuenow",
        "100"
    );

    progressStatus.textContent =
        "Complete";

}


/* ========================================
   DOWNLOAD
======================================== */

function downloadConvertedFile(
    blob,
    filename
) {

    const url =
        URL.createObjectURL(
            blob
        );


    const link =
        document.createElement(
            "a"
        );


    link.href =
        url;

    link.download =
        filename;


    document.body.appendChild(
        link
    );


    link.click();

    link.remove();


    setTimeout(
        () => {

            URL.revokeObjectURL(
                url
            );

        },
        5000
    );

}


/* ========================================
   SUCCESS
======================================== */

function showSuccessScreen(
    blob,
    filename
) {

    convertedBlob =
        blob;

    convertedOutputName =
        filename;

    convertedFileName.textContent =
        filename;

    converterTabs.hidden =
        true;

    uploadArea.hidden =
        true;

    googleArea.hidden =
        true;

    successArea.hidden =
        false;

}


/* ========================================
   LOCAL CONVERT
======================================== */

convertButton.addEventListener(
    "click",
    async () => {

        const file =
            fileInput.files?.[0];


        const validation =
            validateFile(
                file
            );


        if (
            !validation.valid
        ) {

            setStatus(
                validation.message,
                "error"
            );

            return;

        }


        const endpoint =
            conversionMode ===
            "word-to-pdf"
                ?
                "/convert/word-to-pdf"
                :
                "/convert/pdf-to-word";


        const formData =
            new FormData();


        formData.append(
            "file",
            file
        );


        disableAllTabs(
            true
        );

        convertButton.disabled =
            true;

        fileInput.disabled =
            true;

        convertButton.textContent =
            "Converting...";


        startProgress();


        try {

            const response =
                await fetch(
                    endpoint,
                    {
                        method:
                            "POST",

                        body:
                            formData
                    }
                );


            if (
                !response.ok
            ) {

                let message =
                    "Conversion failed.";


                try {

                    const error =
                        await response.json();

                    if (
                        error?.error
                    ) {

                        message =
                            error.error;

                    }

                }

                catch {}


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


            await delay(
                500
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

            progressContainer.hidden =
                true;


            setStatus(
                error.message,
                "error"
            );

        }

        finally {

            disableAllTabs(
                false
            );

            convertButton.disabled =
                false;

            fileInput.disabled =
                false;

            convertButton.textContent =
                "Convert File";

        }

    }
);


/* ========================================
   GOOGLE LIBRARY LOAD
======================================== */

function onGoogleApiLoad() {

    gapi.load(
        "picker",
        () => {

            googlePickerReady =
                true;

            console.log(
                "Google Picker ready"
            );

        }
    );

}


function gisLoaded() {

    googleTokenClient =
        google.accounts.oauth2.initTokenClient({

            client_id:
                GOOGLE_CLIENT_ID,

            scope:
                GOOGLE_SCOPE,

            callback:
                ""

        });


    googleIdentityReady =
        true;


    console.log(
        "Google Identity ready"
    );

}


window.onGoogleApiLoad =
    onGoogleApiLoad;

window.gisLoaded =
    gisLoaded;


/* ========================================
   GOOGLE BUTTONS
======================================== */

chooseGoogleDocButton.addEventListener(
    "click",
    () => {

        requestedPickerType =
            "drive";

        startGooglePicker();

    }
);


uploadGoogleDocButton.addEventListener(
    "click",
    () => {

        requestedPickerType =
            "upload";

        startGooglePicker();

    }
);


/* ========================================
   START GOOGLE PICKER
======================================== */

function startGooglePicker() {

    setGoogleStatus("");


    if (
        !googlePickerReady
    ) {

        setGoogleStatus(
            "Google Picker is still loading. Please wait a moment.",
            "error"
        );

        return;

    }


    if (
        !googleIdentityReady ||
        !googleTokenClient
    ) {

        setGoogleStatus(
            "Google sign-in is still loading. Please wait a moment.",
            "error"
        );

        return;

    }


    requestGoogleAccess();

}


/* ========================================
   GOOGLE AUTH
======================================== */

function requestGoogleAccess() {

    googleTokenClient.callback =
        response => {

            if (
                response.error
            ) {

                setGoogleStatus(
                    "Google authorization failed.",
                    "error"
                );

                return;

            }


            googleAccessToken =
                response.access_token;


            showGooglePicker();

        };


    if (
        googleAccessToken
    ) {

        showGooglePicker();

    }

    else {

        googleTokenClient.requestAccessToken({
            prompt:
                "consent"
        });

    }

}


/* ========================================
   GOOGLE PICKER
======================================== */

function showGooglePicker() {

    const builder =
        new google.picker.PickerBuilder();


    /*
        CHOOSE FROM DRIVE
    */

    if (
        requestedPickerType ===
        "drive"
    ) {

        const driveView =
            new google.picker.DocsView(
                google.picker.ViewId.DOCUMENTS
            );


        if (
            conversionMode ===
            "google-to-pdf"
        ) {

            driveView.setMimeTypes(
                [
                    GOOGLE_DOC_MIME,
                    DOCX_MIME,
                    DOC_MIME
                ].join(",")
            );

        }

        else {

            driveView.setMimeTypes(
                [
                    GOOGLE_DOC_MIME,
                    PDF_MIME
                ].join(",")
            );

        }


        builder.addView(
            driveView
        );

    }


    /*
        UPLOAD FROM COMPUTER
    */

    else {

        const uploadView =
            new google.picker.DocsUploadView();


        builder.addView(
            uploadView
        );

    }


    const picker =
        builder

            .setOAuthToken(
                googleAccessToken
            )

            .setDeveloperKey(
                GOOGLE_API_KEY
            )

            .setAppId(
                GOOGLE_APP_ID
            )

            .setOrigin(
                window.location.origin
            )

            .setTitle(
                requestedPickerType ===
                "upload"
                    ?
                    "Upload from Computer"
                    :
                    "Choose from Google Drive"
            )

            .setCallback(
                googlePickerCallback
            )

            .build();


    picker.setVisible(
        true
    );

}
