from docx.shared import RGBColor

def replace_text_preserving_format(doc, placeholder, replacement):
    """ Replaces {Name} without losing formatting & color """
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if placeholder in run.text:
                text_parts = run.text.split(placeholder)
                run.text = text_parts[0]  # Keep the text before {Name}

                # Create a new run for the replacement with same formatting
                new_run = paragraph.add_run(replacement)
                new_run.bold = run.bold
                new_run.italic = run.italic
                new_run.underline = run.underline
                new_run.font.color.rgb = run.font.color.rgb if run.font.color else RGBColor(0, 0, 0)

                if len(text_parts) > 1:
                    run.text += text_parts[1]  # Append text after {Name}

    # Also replace inside tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if placeholder in run.text:
                            text_parts = run.text.split(placeholder)
                            run.text = text_parts[0]  # Keep text before {Name}

                            new_run = paragraph.add_run(replacement)
                            new_run.bold = run.bold
                            new_run.italic = run.italic
                            new_run.underline = run.underline
                            new_run.font.color.rgb = run.font.color.rgb if run.font.color else RGBColor(0, 0, 0)

                            if len(text_parts) > 1:
                                run.text += text_parts[1]  # Append text after {Name}
