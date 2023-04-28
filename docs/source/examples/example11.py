from pptreport import PowerPointReport

report = PowerPointReport(template="content/template.pptx")

report.add_title_slide(title="How to deal with missing files")
report.add_slide(["content/lion.jpg", "content/not_present*.jpg"],
                 title="With missing_file = 'empty'",
                 missing_file="empty")
report.add_slide(["content/lion.jpg", "content/not_present*.jpg"],
                 title="With missing_file = 'text'",
                 missing_file="text")
report.add_slide(["content/lion.jpg", "content/not_present*.jpg"],
                 title="With missing_file = 'skip'",
                 missing_file="skip")
report.add_slide(["content/lion.jpg", "content/not_present*.jpg"],
                 title="With missing_file = 'skip-slide'",
                 missing_file="skip-slide")   # this slide will not be shown

report.write_config("example11.json")
report.save("example11.pptx", pdf=True)
