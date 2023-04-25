from pptreport import PowerPointReport

report = PowerPointReport(template="content/template.pptx")

report.add_title_slide(title="Add text from textfiles")
report.add_slide(content=["content/*_fish*", "content/fish_description.txt"],
                 title="Fish with text from file")

report.write_config("example12.json")
report.save("example12.pptx", pdf=True)
