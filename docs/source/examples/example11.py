from pptreport import PowerPointReport

report = PowerPointReport()

report.add_title_slide(title="How to deal with missing files")
report.add_slide(["content/lion.jpg", "content/not_present*.jpg"],
                 title="With missing_file = 'empty'",
                 missing_file="empty")
report.add_slide(["content/lion.jpg", "content/not_present*.jpg"],
                 title="With missing_file = 'skip'",
                 missing_file="skip")

report.write_config("example11.json")
report.save("example11.pptx", pdf=True)
