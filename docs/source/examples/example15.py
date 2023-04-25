from pptreport import PowerPointReport

report = PowerPointReport(template="content/template.pptx")

report.add_title_slide(title="Content from PDFs")
report.add_slide(content=["content/*_fish*", "content/chips.pdf", "The chips came from a .pdf."],
                 title="Fish and chips")
report.add_slide(content=["content/pdfs/multidogs.pdf"],
                 title="With pdf_pages='all', all pages are used",
                 pdf_pages="all")
report.add_slide(content=["content/pdfs/multidogs.pdf"],
                 title="With pdf_pages=[1,3]",
                 pdf_pages=[1, 3])
report.add_slide(content=["content/pdfs/multidogs.pdf"],
                 title="Split from pdf with split=True",
                 split=True)

# Grouping is also possible with pdfs, but only one page is allowed
report.add_slide(title="Grouping with pdf",
                 grouped_content=["content/pdfs/(\w+)_blue.pdf",
                                  "content/pdfs/(.*)_yellow.pdf",
                                  "content/pdfs/([a-z\_]+)_red.pdf"],
                 n_columns=3,
                 show_filename=True,
                 pdf_pages=1)

report.write_config("example15.json")
report.save("example15.pptx", pdf=True)
