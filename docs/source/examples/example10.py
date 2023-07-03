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
report.add_slide(["content/lion-not-present.jpg", "content/not_present*.jpg"],
                 title="With missing_file='text' and empty_slide='keep'",
                 missing_file="text")
report.add_slide(["content/lion-not-present.jpg", "content/not_present*.jpg"],
                 title="With missing_file='text' and empty_slide='skip'",
                 missing_file="text",
                 empty_slide="skip")  # this slide will not be shown
report.add_slide(["subheader 1", "subheader 2", "content/lion-not-present.jpg", "content/not_present*.jpg"],
                 title="With missing_file='text' and empty_slide='skip'",
                 missing_file="text",
                 empty_slide="skip")  # this slide will also not be shown even if headers are present

report.write_config("example10.json")
report.save("example10.pptx", pdf=True)
