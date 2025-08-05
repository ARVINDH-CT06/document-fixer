def analyze_document(doc):
    issues = []

    # Check for unformatted paragraphs
    for para in doc.paragraphs:
        if para.text.strip() and para.alignment != 3:  # 3 = Justify
            issues.append("Some paragraphs were not justified and needed fixing.")
            break

    # Check for tables
    if doc.tables:
        issues.append(f"Found {len(doc.tables)} table(s) - formatting applied.")

    if not issues:
        issues.append("Document already followed basic structure.")
    return issues
