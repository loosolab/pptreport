from pptreport import PowerPointReport

report = PowerPointReport(template="content/template.pptx")

report.add_title_slide(title="Show borders around boxes")
report.add_slide(content=["content/*_fish*", "content/chips.pdf", "The chips came from a .pdf."],
                 title="show_borders=True",
                 show_borders=True)
report.add_slide(content=["content/*_fish*", "content/chips.pdf", "The chips came from a .pdf."],
                 title="show_borders=False (default)",
                 show_borders=False)

report.write_config("example16.json")
report.save("example16.pptx", pdf=True)
