import docx

doc = docx.Document("1. Huong dan trinh bay DATN.docx")
print("Total sections:", len(doc.sections))
for i, sect in enumerate(doc.sections):
    print(f"Section {i}:")
    footer = sect.footer
    print("  Linked:", footer.is_linked_to_previous)
    xml = footer._element.xml
    print("  Has PAGE:", 'PAGE' in xml)
