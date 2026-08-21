# =========================================
# PDF TO WORD
# USING FREE LOCAL PDF2WORD CONVERTER
# =========================================

@app.route(
    "/convert/pdf-to-word",
    methods=["POST"]
)
def pdf_to_word():

    if "file" not in request.files:

        return jsonify({
            "error":
            "No file uploaded."
        }), 400


    uploaded_file = (
        request.files["file"]
    )


    if uploaded_file.filename == "":

        return jsonify({
            "error":
            "No file selected."
        }), 400


    filename = secure_filename(
        uploaded_file.filename
    )


    if not filename.lower().endswith(
        ".pdf"
    ):

        return jsonify({
            "error":
            "Please upload a PDF file."
        }), 400


    with tempfile.TemporaryDirectory() as temp_dir:

        input_path = os.path.join(
            temp_dir,
            filename
        )


        uploaded_file.save(
            input_path
        )


        base_name = (
            os.path.splitext(
                filename
            )[0]
        )


        output_name = (
            base_name +
            ".docx"
        )


        output_path = os.path.join(
            temp_dir,
            output_name
        )


        try:

            result = subprocess.run(
                [
                    "pdf2word",
                    "convert",
                    input_path,
                    output_path
                ],

                check=True,

                timeout=180,

                capture_output=True,

                text=True
            )


            if result.stdout:

                print(
                    "pdf2word output:",
                    result.stdout
                )


            if result.stderr:

                print(
                    "pdf2word warnings:",
                    result.stderr
                )


        except subprocess.TimeoutExpired:

            return jsonify({
                "error":
                "The PDF conversion took too long."
            }), 504


        except subprocess.CalledProcessError as error:

            print(
                "pdf2word error:",
                error.stderr
            )


            return jsonify({
                "error":
                "The PDF could not be converted to Word."
            }), 500


        except Exception as error:

            print(
                "PDF to Word error:",
                repr(error)
            )


            return jsonify({
                "error":
                "The PDF could not be converted to Word."
            }), 500


        if not os.path.exists(
            output_path
        ):

            return jsonify({
                "error":
                "The Word document could not be created."
            }), 500


        if os.path.getsize(
            output_path
        ) == 0:

            return jsonify({
                "error":
                "The Word document was created but was empty."
            }), 500


        return send_file(
            output_path,

            as_attachment=True,

            download_name=(
                output_name
            ),

            mimetype=(
                "application/vnd."
                "openxmlformats-officedocument."
                "wordprocessingml.document"
            )
        )
