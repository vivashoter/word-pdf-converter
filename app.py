def add_la350_top(
    document,
    values
):

    outer = document.add_table(
        rows=1,
        cols=2
    )

    outer.alignment = (
        WD_TABLE_ALIGNMENT.CENTER
    )

    outer.autofit = False

    set_table_fixed_layout(
        outer
    )

    # Keep the proportions from test (18).
    # These are now structurally much closer to the PDF.
    widths = [
        4.85,
        2.45
    ]

    set_table_column_widths(
        outer,
        widths
    )

    left = outer.cell(
        0,
        0
    )

    right = outer.cell(
        0,
        1
    )

    remove_cell_borders(
        left
    )

    remove_cell_borders(
        right
    )

    set_cell_margins(
        left,
        top=0,
        bottom=0,
        start=0,
        end=10
    )

    set_cell_margins(
        right,
        top=0,
        bottom=0,
        start=0,
        end=0
    )

    left.vertical_alignment = (
        WD_CELL_VERTICAL_ALIGNMENT.TOP
    )

    right.vertical_alignment = (
        WD_CELL_VERTICAL_ALIGNMENT.TOP
    )

    clear_cell(left)
    clear_cell(right)

    # ========================================================
    # LEFT HEADER
    # ========================================================

    header = left.add_table(
        rows=1,
        cols=2
    )

    header.autofit = False

    set_table_fixed_layout(
        header
    )

    # Keep the good horizontal proportions from test (18).
    header_widths = [
        1.45,
        3.30
    ]

    set_table_column_widths(
        header,
        header_widths
    )

    # --------------------------------------------------------
    # LA-350 BLACK BOX
    # --------------------------------------------------------

    black = header.cell(
        0,
        0
    )

    set_cell_shading(
        black,
        "000000"
    )

    remove_cell_borders(
        black
    )

    set_cell_margins(
        black,
        top=7,
        bottom=7,
        start=10,
        end=10
    )

    paragraph = clear_cell(
        black
    )

    paragraph.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    add_run(
        paragraph,
        "LA-350",
        size=14.5,
        bold=True,
        color=(
            255,
            255,
            255
        )
    )

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    title = header.cell(
        0,
        1
    )

    remove_cell_borders(
        title
    )

    set_cell_margins(
        title,
        top=2,
        bottom=1,
        start=10,
        end=0
    )

    paragraph = clear_cell(
        title
    )

    add_run(
        paragraph,
        "Notice of Available Language\n"
        "Assistance—Service Provider",
        size=12.5,
        bold=True
    )

    # Test (18): 42
    # New: 38
    set_row_height(
        header.rows[0],
        38,
        exact=True
    )

    keep_table_row_together(
        header.rows[0]
    )

    # ========================================================
    # BLACK LINE UNDER HEADER
    # ========================================================

    line_table = left.add_table(
        rows=1,
        cols=1
    )

    line_cell = line_table.cell(
        0,
        0
    )

    remove_cell_borders(
        line_cell
    )

    set_cell_margins(
        line_cell,
        top=0,
        bottom=0,
        start=0,
        end=0
    )

    paragraph = clear_cell(
        line_cell
    )

    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)

    run = paragraph.add_run(
        " "
    )

    run.font.size = Pt(1)

    border = {
        "val": "single",
        "sz": 14,
        "color": "000000"
    }

    set_cell_border(
        line_cell,
        bottom=border
    )

    set_row_height(
        line_table.rows[0],
        3,
        exact=True
    )

    # ========================================================
    # INSTRUCTIONS
    # ========================================================

    instruction_table = left.add_table(
        rows=1,
        cols=1
    )

    instruction_cell = instruction_table.cell(
        0,
        0
    )

    remove_cell_borders(
        instruction_cell
    )

    set_cell_margins(
        instruction_cell,
        top=1,
        bottom=0,
        start=0,
        end=5
    )

    paragraph = clear_cell(
        instruction_cell
    )

    add_run(
        paragraph,
        "Use this form to:",
        size=7.2,
        bold=True
    )

    paragraph = instruction_cell.add_paragraph()

    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 0.86

    add_run(
        paragraph,
        "• Tell the court that you are a service provider, "
        "program, or professional offering language assistance "
        "with services that may be ordered by a court; and",
        size=6.5
    )

    paragraph = instruction_cell.add_paragraph()

    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 0.86

    add_run(
        paragraph,
        "• Provide information about the services you provide, "
        "the languages and types of language assistance available, "
        "and your service area.",
        size=6.5
    )

    # Test (18): 54
    # New: 42
    set_row_height(
        instruction_table.rows[0],
        42,
        exact=True
    )

    keep_table_row_together(
        instruction_table.rows[0]
    )

    # ========================================================
    # SECTION 1
    # ========================================================

    section1 = left.add_table(
        rows=1,
        cols=2
    )

    section1.autofit = False

    set_table_fixed_layout(
        section1
    )

    section_widths = [
        0.36,
        4.39
    ]

    set_table_column_widths(
        section1,
        section_widths
    )

    number_cell = section1.cell(
        0,
        0
    )

    content_cell = section1.cell(
        0,
        1
    )

    remove_cell_borders(
        number_cell
    )

    remove_cell_borders(
        content_cell
    )

    set_cell_margins(
        number_cell,
        top=1,
        bottom=0,
        start=0,
        end=0
    )

    set_cell_margins(
        content_cell,
        top=1,
        bottom=0,
        start=0,
        end=2
    )

    paragraph = clear_cell(
        number_cell
    )

    add_section_number(
        paragraph,
        1
    )

    paragraph = clear_cell(
        content_cell
    )

    paragraph.paragraph_format.line_spacing = 0.90

    add_run(
        paragraph,
        "This form should be filed with the court by January 31 "
        "of each year to indicate services that will be provided "
        "during the calendar year. You may also submit this form "
        "to let the court know your services have changed.",
        size=6.5
    )

    paragraph = content_cell.add_paragraph()

    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 0.90

    add_run(
        paragraph,
        "The information in this form describes services "
        "available during calendar year: ",
        size=6.5
    )

    add_run(
        paragraph,
        (
            values["calendar_year"]
            if values["calendar_year"]
            else "____________________________"
        ),
        size=6.5
    )

    # Test (18): 70
    # New: 49
    set_row_height(
        section1.rows[0],
        49,
        exact=True
    )

    keep_table_row_together(
        section1.rows[0]
    )

    # ========================================================
    # SECTION 2
    # ========================================================

    section2 = left.add_table(
        rows=4,
        cols=2
    )

    section2.autofit = False

    set_table_fixed_layout(
        section2
    )

    set_table_column_widths(
        section2,
        section_widths
    )

    for row_index in range(
        4
    ):

        remove_cell_borders(
            section2.cell(
                row_index,
                0
            )
        )

        remove_cell_borders(
            section2.cell(
                row_index,
                1
            )
        )

        set_cell_margins(
            section2.cell(
                row_index,
                0
            ),
            top=0,
            bottom=0,
            start=0,
            end=0
        )

        set_cell_margins(
            section2.cell(
                row_index,
                1
            ),
            top=0,
            bottom=0,
            start=0,
            end=2
        )

    paragraph = clear_cell(
        section2.cell(
            0,
            0
        )
    )

    add_section_number(
        paragraph,
        2
    )

    # --------------------------------------------------------
    # PROVIDER
    # --------------------------------------------------------

    paragraph = clear_cell(
        section2.cell(
            0,
            1
        )
    )

    add_form_line(
        paragraph,
        "Name of service provider:",
        values["provider"],
        width_chars=46,
        size=6.7
    )

    # --------------------------------------------------------
    # ADDRESS
    # --------------------------------------------------------

    paragraph = clear_cell(
        section2.cell(
            1,
            1
        )
    )

    add_form_line(
        paragraph,
        "Address:",
        values["address"],
        width_chars=65,
        size=6.7
    )

    # --------------------------------------------------------
    # TELEPHONE / WEB
    # --------------------------------------------------------

    paragraph = clear_cell(
        section2.cell(
            2,
            1
        )
    )

    add_form_line(
        paragraph,
        "Telephone:",
        values["telephone"],
        width_chars=18,
        size=6.7
    )

    add_run(
        paragraph,
        "   ",
        size=6.7
    )

    add_form_line(
        paragraph,
        "Web address:",
        values["web"],
        width_chars=29,
        size=6.7
    )

    # --------------------------------------------------------
    # CONTACT / EMAIL
    # --------------------------------------------------------

    paragraph = clear_cell(
        section2.cell(
            3,
            1
        )
    )

    add_form_line(
        paragraph,
        "Contact name:",
        values["contact_name"],
        width_chars=25,
        size=6.7
    )

    add_run(
        paragraph,
        "   ",
        size=6.7
    )

    add_form_line(
        paragraph,
        "E-mail:",
        values["email"],
        width_chars=29,
        size=6.7
    )

    # Test (18): 21 pt each
    # New: 17 pt each
    for row_index in range(
        4
    ):

        set_row_height(
            section2.rows[
                row_index
            ],
            17,
            exact=True
        )

        keep_table_row_together(
            section2.rows[
                row_index
            ]
        )

    # ========================================================
    # RIGHT SIDE
    # ========================================================

    right_table = right.add_table(
        rows=2,
        cols=1
    )

    right_table.autofit = False

    set_table_fixed_layout(
        right_table
    )

    set_table_column_widths(
        right_table,
        [
            2.45
        ]
    )

    # ========================================================
    # CLERK STAMP BOX
    # ========================================================

    clerk = right_table.cell(
        0,
        0
    )

    set_cell_margins(
        clerk,
        top=3,
        bottom=3,
        start=8,
        end=8
    )

    paragraph = clear_cell(
        clerk
    )

    add_run(
        paragraph,
        "Clerk stamps date here when form is received.",
        size=5.4,
        italic=True
    )

    border = {
        "val": "single",
        "sz": 6,
        "color": "000000"
    }

    set_cell_border(
        clerk,
        top=border,
        bottom=border,
        left=border,
        right=border
    )

    # Test (18): 139
    # New: 110
    set_row_height(
        right_table.rows[0],
        110,
        exact=True
    )

    keep_table_row_together(
        right_table.rows[0]
    )

    # ========================================================
    # COURT INFORMATION BOX
    # ========================================================

    court = right_table.cell(
        1,
        0
    )

    set_cell_margins(
        court,
        top=3,
        bottom=3,
        start=8,
        end=8
    )

    paragraph = clear_cell(
        court
    )

    add_run(
        paragraph,
        "Fill in court name and address:",
        size=5.4,
        italic=True
    )

    paragraph = court.add_paragraph()

    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)

    add_run(
        paragraph,
        "Superior Court of California, County of",
        size=6.2,
        bold=True
    )

    paragraph = court.add_paragraph()

    paragraph.paragraph_format.space_before = Pt(1)
    paragraph.paragraph_format.space_after = Pt(0)

    add_run(
        paragraph,
        (
            values["court"]
            if values["court"]
            else " "
        ),
        size=6.5
    )

    set_cell_border(
        court,
        top=border,
        bottom=border,
        left=border,
        right=border
    )

    # Test (18): 91
    # New: 78
    set_row_height(
        right_table.rows[1],
        78,
        exact=True
    )

    keep_table_row_together(
        right_table.rows[1]
    )

    keep_table_row_together(
        outer.rows[0]
    )

    return outer
