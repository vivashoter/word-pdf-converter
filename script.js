/* ========================================
   GOOGLE CONFIGURATION
======================================== */

const GOOGLE_CLIENT_ID =
    "239816509439-uhrb687i85r1susl9jmbn5it4afcdebe.apps.googleusercontent.com";

const GOOGLE_API_KEY =
    "AIzaSyCAv-JotC9-W5JPMrbQ4T-25GYZ03d-kB4";

const GOOGLE_APP_ID =
    "239816509439";

const GOOGLE_SCOPE =
    "https://www.googleapis.com/auth/drive.file";


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

const uploadTitle =
    document.getElementById("uploadTitle");

const uploadDescription =
    document.getElementById("uploadDescription");

const uploadArea =
    document.getElementById("uploadArea");

const googleArea =
    document.getElementById("googleArea");

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
   STATE
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

let googleProgressTimer =
    null;

let googleProgressValue =
    0;


/* GOOGLE STATE */

let googleTokenClient =
    null;

let googleAccessToken =
    null;

let googlePickerReady =
    false;

let googleIdentityReady =
    false;

let selectedGoogleDocId =
    null;

let selectedGoogleDocName =
    null;

let selectedGoogleDocMimeType =
    null;


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
   MODE HELPERS
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


function setMode(mode) {

    conversionMode =
        mode;

    clearActiveTabs();

    successArea.hidden =
        true;

    converterTabs.hidden =
        false;


    if (
        mode ===
        "word-to-pdf"
    ) {

        wordToPdfButton.classList.add(
            "active"
        );

        uploadArea.hidden =
            false;

        googleArea.hidden =
            true;

        fileInput.accept =
            ".doc,.docx";

        uploadArea.className =
            "upload-area word-mode";

        resetFileSelection();

    }


    else if (
        mode ===
        "pdf-to-word"
    ) {

        pdfToWordButton.classList.add(
            "active"
        );

        uploadArea.hidden =
            false;

        googleArea.hidden =
            true;

        fileInput.accept =
            ".pdf";

        uploadArea.className =
            "upload-area pdf-mode";

        resetFileSelection();

    }


    else if (
        mode ===
        "google-to-pdf"
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
        mode ===
        "google-to-word"
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
            "Choose a Google Doc from Drive or upload a Word document from your computer.";

        resetGoogleSelection();

    }

}


/* ========================================
   RESET LOCAL UPLOAD TEXT
======================================== */

function resetUploadText() {

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
   LOCAL PROGRESS
======================================== */

function resetProgress() {

    if (
        progressTimer
    ) {

        clearInterval(
            progressTimer
        );

        progressTimer =
            null;

    }

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

    progressNote.textContent =
        "Please keep this page open while your document is converted.";

}


function startProgress() {

    resetProgress();

    progressContainer.hidden =
        false;

    progressValue =
        6;

    progressBar.style.width =
        progressValue +
        "%";

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

                let target;


                if (
                    elapsedSeconds <
                    2
                ) {

                    target =
                        18;

                    progressStatus.textContent =
                        "Uploading";

                }


                else if (
                    elapsedSeconds <
                    5
                ) {

                    target =
                        38;

                    progressStatus.textContent =
                        "Reading document";

                }


                else if (
                    elapsedSeconds <
                    10
                ) {

                    target =
                        58;

                    progressStatus.textContent =
                        "Converting";

                }


                else if (
                    elapsedSeconds <
                    20
                ) {

                    target =
                        74;

                    progressStatus.textContent =
                        "Converting";

                }


                else if (
                    elapsedSeconds <
                    35
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


function completeProgress() {

    if (
        progressTimer
    ) {

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


/* ========================================
   RESET LOCAL FILE
======================================== */

function resetFileSelection() {

    fileInput.value =
        "";

    progressContainer.hidden =
        true;

    resetProgress();

    setStatus("");

    resetUploadText();

}


/* ========================================
   TAB EVENTS
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
   VALIDATE LOCAL FILE
======================================== */

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
            !fileName.endsWith(
                ".doc"
            ) &&
            !fileName.endsWith(
                ".docx"
            )
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
            !fileName.endsWith(
                ".pdf"
            )
        ) {

            return {
                valid: false,

                message:
                    "Please choose a PDF file."
            };

        }

    }


    const maximumFileSize =
        25 *
        1024 *
        1024;


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


/* ========================================
   SELECT LOCAL FILE
======================================== */

function showSelectedFile(file) {

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

    progressContainer.hidden =
        true;

    setStatus("");

    return true;

}


fileInput.addEventListener(
    "change",
    () => {

        if (
            !fileInput.files ||
            fileInput.files.length ===
            0
        ) {

            return;

        }


        showSelectedFile(
            fileInput.files[0]
        );

    }
);


/* ========================================
   DRAG AND DROP
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
            event.dataTransfer.files[0];


        if (!file) {

            return;

        }


        if (
            !showSelectedFile(
                file
            )
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


/* ========================================
   DOWNLOAD
======================================== */

function downloadConvertedFile(
    blob,
    outputName
) {

    const downloadUrl =
        URL.createObjectURL(
            blob
        );


    const link =
        document.createElement(
            "a"
        );


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


/* ========================================
   SUCCESS
======================================== */

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

    googleArea.hidden =
        true;

    successArea.hidden =
        false;

}


/* ========================================
   DOWNLOAD AGAIN
======================================== */

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


/* ========================================
   CONVERT ANOTHER
======================================== */

convertAnotherButton.addEventListener(
    "click",
    () => {

        convertedBlob =
            null;

        convertedOutputName =
            "";

        selectedGoogleDocId =
            null;

        selectedGoogleDocName =
            null;

        selectedGoogleDocMimeType =
            null;

        successArea.hidden =
            true;

        converterTabs.hidden =
            false;

        setMode(
            conversionMode
        );

    }
);


/* ========================================
   LOCAL CONVERSION
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


        const formData =
            new FormData();


        formData.append(
            "file",
            file
        );


        const endpoint =
            conversionMode ===
            "word-to-pdf"
                ?
                "/convert/word-to-pdf"
                :
                "/convert/pdf-to-word";


        disableAllTabs(
            true
        );

        fileInput.disabled =
            true;

        convertButton.disabled =
            true;

        convertButton.textContent =
            "Converting...";


        startProgress();

        setStatus("");


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


            await delay(
                650
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

            disableAllTabs(
                false
            );


            fileInput.disabled =
                false;

            convertButton.disabled =
                false;

            convertButton.textContent =
                "Convert File";

        }

    }
);


/* ========================================
   GOOGLE API LOADING
======================================== */

function onGoogleApiLoad() {

    gapi.load(
        "picker",
        () => {

            googlePickerReady =
                true;

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

}


window.onGoogleApiLoad =
    onGoogleApiLoad;

window.gisLoaded =
    gisLoaded;


/* ========================================
   CHOOSE GOOGLE DOC
======================================== */

chooseGoogleDocButton.addEventListener(
    "click",
    () => {

        if (
            GOOGLE_CLIENT_ID.includes(
                "PASTE_"
            ) ||
            GOOGLE_API_KEY.includes(
                "PASTE_"
            )
        ) {

            setGoogleStatus(
                "Google credentials have not been added yet.",
                "error"
            );

            return;

        }


        if (
            !googlePickerReady ||
            !googleIdentityReady
        ) {

            setGoogleStatus(
                "Google Drive is still loading. Please try again in a moment.",
                "error"
            );

            return;

        }


        requestGoogleAccess();

    }
);


/* ========================================
   GOOGLE ACCESS
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

    /*
        Existing Google Docs in Drive
    */

    const googleDocsView =
        new google.picker.DocsView(
            google.picker.ViewId.DOCUMENTS
        );


    googleDocsView.setMimeTypes(

        [
            GOOGLE_DOC_MIME,
            DOCX_MIME,
            DOC_MIME
        ].join(",")

    );


    /*
        Upload from computer
    */

    const uploadView =
        new google.picker.DocsUploadView();


    const picker =
        new google.picker.PickerBuilder()

            .addView(
                googleDocsView
            )

            .addView(
                uploadView
            )

            .setSelectableMimeTypes(

                [
                    GOOGLE_DOC_MIME,
                    DOCX_MIME,
                    DOC_MIME
                ].join(",")

            )

            .setOAuthToken(
                googleAccessToken
            )

            .setDeveloperKey(
                GOOGLE_API_KEY
            )

            .setAppId(
                GOOGLE_APP_ID
            )

            .setTitle(
                "Choose or Upload a Document"
            )

            .setCallback(
                googlePickerCallback
            )

            .build();


    picker.setVisible(
        true
    );

}


/* ========================================
   PICKER CALLBACK
======================================== */

function googlePickerCallback(
    data
) {

    const action =
        data[
            google.picker.Response.ACTION
        ];


    if (
        action ===
        google.picker.Action.PICKED
    ) {

        const document =
            data[
                google.picker.Response.DOCUMENTS
            ][0];


        selectedGoogleDocId =
            document[
                google.picker.Document.ID
            ];


        selectedGoogleDocName =
            document[
                google.picker.Document.NAME
            ] ||
            "Document";


        selectedGoogleDocMimeType =
            document[
                google.picker.Document.MIME_TYPE
            ] ||
            null;


        selectedGoogleFileName.textContent =
            selectedGoogleDocName;


        selectedGoogleFile.hidden =
            false;


        convertGoogleButton.hidden =
            false;


        if (
            selectedGoogleDocMimeType ===
            GOOGLE_DOC_MIME
        ) {

            setGoogleStatus(
                "Google Doc selected and ready to convert.",
                "success"
            );

        }


        else if (
            selectedGoogleDocMimeType ===
            DOCX_MIME ||
            selectedGoogleDocMimeType ===
            DOC_MIME
        ) {

            setGoogleStatus(
                "Word document selected and ready to convert.",
                "success"
            );

        }


        else {

            setGoogleStatus(
                "This file type is not supported.",
                "error"
            );

        }

    }

}


/* ========================================
   GOOGLE PROGRESS
======================================== */

function startGoogleProgress() {

    resetGoogleProgress();


    googleProgressContainer.hidden =
        false;


    googleProgressValue =
        8;


    googleProgressBar.style.width =
        "8%";


    googleProgressStatus.textContent =
        "Preparing";


    googleProgressTimer =
        setInterval(
            () => {

                if (
                    googleProgressValue <
                    90
                ) {

                    googleProgressValue +=
                        Math.max(
                            1,
                            (
                                90 -
                                googleProgressValue
                            ) *
                            0.08
                        );


                    googleProgressBar.style.width =
                        googleProgressValue +
                        "%";


                    googleProgressTrack.setAttribute(
                        "aria-valuenow",
                        Math.round(
                            googleProgressValue
                        )
                    );

                }


                if (
                    googleProgressValue >
                    35
                ) {

                    googleProgressStatus.textContent =
                        "Converting";

                }


                if (
                    googleProgressValue >
                    75
                ) {

                    googleProgressStatus.textContent =
                        "Finishing";

                }

            },
            400
        );

}


function completeGoogleProgress() {

    if (
        googleProgressTimer
    ) {

        clearInterval(
            googleProgressTimer
        );

        googleProgressTimer =
            null;

    }


    googleProgressValue =
        100;


    googleProgressBar.style.width =
        "100%";


    googleProgressTrack.setAttribute(
        "aria-valuenow",
        "100"
    );


    googleProgressStatus.textContent =
        "Complete";

}


function resetGoogleProgress() {

    if (
        googleProgressTimer
    ) {

        clearInterval(
            googleProgressTimer
        );

        googleProgressTimer =
            null;

    }


    googleProgressValue =
        0;


    googleProgressBar.style.width =
        "0%";


    googleProgressTrack.setAttribute(
        "aria-valuenow",
        "0"
    );

}


/* ========================================
   RESET GOOGLE
======================================== */

function resetGoogleSelection() {

    selectedGoogleDocId =
        null;

    selectedGoogleDocName =
        null;

    selectedGoogleDocMimeType =
        null;

    selectedGoogleFile.hidden =
        true;

    convertGoogleButton.hidden =
        true;

    googleProgressContainer.hidden =
        true;

    resetGoogleProgress();

    setGoogleStatus("");

}


/* ========================================
   DOWNLOAD DRIVE BLOB FILE
======================================== */

async function downloadDriveBlobFile(
    fileId
) {

    const url =
        "https://www.googleapis.com/drive/v3/files/" +
        encodeURIComponent(
            fileId
        ) +
        "?alt=media";


    const response =
        await fetch(
            url,
            {
                headers: {

                    Authorization:
                        "Bearer " +
                        googleAccessToken

                }
            }
        );


    if (
        !response.ok
    ) {

        throw new Error(
            "Could not download the uploaded document from Google Drive."
        );

    }


    return await response.blob();

}


/* ========================================
   GOOGLE DOC EXPORT
======================================== */

async function exportNativeGoogleDoc() {

    let mimeType;

    let extension;


    if (
        conversionMode ===
        "google-to-pdf"
    ) {

        mimeType =
            "application/pdf";

        extension =
            ".pdf";

        googleProgressLabel.textContent =
            "Exporting Google Doc to PDF…";

    }


    else {

        mimeType =
            DOCX_MIME;

        extension =
            ".docx";

        googleProgressLabel.textContent =
            "Exporting Google Doc to Word…";

    }


    const exportUrl =
        "https://www.googleapis.com/drive/v3/files/" +
        encodeURIComponent(
            selectedGoogleDocId
        ) +
        "/export?mimeType=" +
        encodeURIComponent(
            mimeType
        );


    const response =
        await fetch(
            exportUrl,
            {

                headers: {

                    Authorization:
                        "Bearer " +
                        googleAccessToken

                }

            }
        );


    if (
        !response.ok
    ) {

        let errorMessage =
            "Google could not export this document.";


        try {

            const errorData =
                await response.json();


            if (
                errorData?.error?.message
            ) {

                errorMessage =
                    errorData.error.message;

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


    return {

        blob:
            blob,

        outputName:
            removeExtension(
                selectedGoogleDocName
            ) +
            extension

    };

}


/* ========================================
   HANDLE UPLOADED WORD FILE
======================================== */

async function convertUploadedWordFile() {

    googleProgressLabel.textContent =
        "Processing uploaded Word document…";


    const driveBlob =
        await downloadDriveBlobFile(
            selectedGoogleDocId
        );


    /*
        Google Doc → Word mode

        The uploaded file is already
        a Word document, so download it.
    */

    if (
        conversionMode ===
        "google-to-word"
    ) {

        let extension =
            ".docx";


        if (
            selectedGoogleDocMimeType ===
            DOC_MIME
        ) {

            extension =
                ".doc";

        }


        return {

            blob:
                driveBlob,

            outputName:
                removeExtension(
                    selectedGoogleDocName
                ) +
                extension

        };

    }


    /*
        Google Doc → PDF mode

        Send uploaded Word file through
        the working Word-to-PDF backend.
    */

    const formData =
        new FormData();


    formData.append(
        "file",
        driveBlob,
        selectedGoogleDocName
    );


    const response =
        await fetch(
            "/convert/word-to-pdf",
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

        let errorMessage =
            "The uploaded Word document could not be converted to PDF.";


        try {

            const errorData =
                await response.json();


            if (
                errorData?.error
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


    const pdfBlob =
        await response.blob();


    return {

        blob:
            pdfBlob,

        outputName:
            removeExtension(
                selectedGoogleDocName
            ) +
            ".pdf"

    };

}


/* ========================================
   GOOGLE CONVERT BUTTON
======================================== */

convertGoogleButton.addEventListener(
    "click",
    async () => {

        if (
            !selectedGoogleDocId
        ) {

            setGoogleStatus(
                "Please choose or upload a document first.",
                "error"
            );

            return;

        }


        if (
            !googleAccessToken
        ) {

            setGoogleStatus(
                "Google authorization expired. Please choose the document again.",
                "error"
            );

            return;

        }


        disableAllTabs(
            true
        );


        chooseGoogleDocButton.disabled =
            true;


        convertGoogleButton.disabled =
            true;


        convertGoogleButton.textContent =
            "Converting...";


        setGoogleStatus("");


        startGoogleProgress();


        try {

            let result;


            /*
                Native Google Doc
            */

            if (
                selectedGoogleDocMimeType ===
                GOOGLE_DOC_MIME
            ) {

                result =
                    await exportNativeGoogleDoc();

            }


            /*
                Word file uploaded through Picker
            */

            else if (
                selectedGoogleDocMimeType ===
                DOCX_MIME ||
                selectedGoogleDocMimeType ===
                DOC_MIME
            ) {

                result =
                    await convertUploadedWordFile();

            }


            else {

                throw new Error(
                    "Please select a Google Doc, DOC, or DOCX file."
                );

            }


            completeGoogleProgress();


            await delay(
                500
            );


            downloadConvertedFile(
                result.blob,
                result.outputName
            );


            showSuccessScreen(
                result.blob,
                result.outputName
            );

        }


        catch (error) {

            resetGoogleProgress();


            googleProgressContainer.hidden =
                true;


            setGoogleStatus(
                error.message ||
                "Document conversion failed.",
                "error"
            );

        }


        finally {

            disableAllTabs(
                false
            );


            chooseGoogleDocButton.disabled =
                false;


            convertGoogleButton.disabled =
                false;


            convertGoogleButton.textContent =
                "Convert Google Doc";

        }

    }
);


/* ========================================
   UTILITIES
======================================== */

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


function removeExtension(
    name
) {

    return name.replace(
        /\.[^/.]+$/,
        ""
    );

}


function delay(
    milliseconds
) {

    return new Promise(
        resolve =>
            setTimeout(
                resolve,
                milliseconds
            )
    );

}


/* ========================================
   INITIAL STATE
======================================== */

setMode(
    "word-to-pdf"
);
