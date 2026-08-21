def flatten_pdf_form(
    input_path,
    output_path
):

    result = subprocess.run(
        [
            "qpdf",

            input_path,

            output_path,

            "--generate-appearances",

            "--flatten-annotations=all",

            "--remove-acroform"
        ],

        check=True,
        timeout=120,
        capture_output=True,
        text=True
    )


    if result.stdout:
        print(
            "QPDF output:",
            result.stdout
        )


    if result.stderr:
        print(
            "QPDF warnings:",
            result.stderr
        )


    if not os.path.exists(output_path):
        return False


    if os.path.getsize(output_path) == 0:
        return False


    print(
        "QPDF flattened annotations "
        "and removed AcroForm."
    )


    return True
