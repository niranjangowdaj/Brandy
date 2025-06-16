def pptx_compliance_check_with_rules(pptx_file, rules, add_copyright, copyright_type):
    prs = Presentation(pptx_file)
    issues = []
    slide_comments = {}
    
    if add_copyright:
        copyright_text = "© SAP SE or an SAP affiliate company. All rights reserved. Internal Use Only." if copyright_type == "Internal" else "© SAP SE or an SAP affiliate company. All rights reserved. Public Use."
        for slide in prs.slides:
            add_footer_to_slide(slide, copyright_text)
    
    for slide_idx, slide in enumerate(prs.slides, 1):
        slide_issue_comments = []
        for shape_idx, shape in enumerate(slide.shapes, 1):
            if not shape.has_text_frame:
                continue
            text = shape.text
            if len(text) > 0:
                # Check for SAP branding
                if "SAP" not in text and slide_idx == 1:
                    issues.append(f"Slide {slide_idx}: Missing SAP branding on title slide")
                    slide_issue_comments.append("Missing SAP branding on title slide")
                    add_red_border(shape)
                
                # Check for font compliance
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        # Check font name
                        if run.font.name and run.font.name not in ["SAP 72", "Arial"]:
                            issues.append(f"Slide {slide_idx}: Non-compliant font '{run.font.name}' used")
                            slide_issue_comments.append(f"Non-compliant font '{run.font.name}' used")
                            add_red_border(shape)
                        
                        # Check font size
                        if run.font.size and run.font.size.pt < 10:
                            issues.append(f"Slide {slide_idx}: Font size too small ({run.font.size.pt}pt) - must be at least 10pt")
                            slide_issue_comments.append(f"Font size too small ({run.font.size.pt}pt) - must be at least 10pt")
                            add_red_border(shape)
        
        # Add speaker notes
        notes_slide = slide.notes_slide
        notes_text_frame = notes_slide.notes_text_frame
        if slide_issue_comments:
            notes_text_frame.text = f"Slide {slide_idx} compliance issues:\n" + "\n".join(slide_issue_comments)
        else:
            notes_text_frame.text = f"Slide {slide_idx}: All elements compliant."
    
    add_summary_slide(prs, issues)
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return issues, output 