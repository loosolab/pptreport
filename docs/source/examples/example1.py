from pptreport import PowerPointReport

report = PowerPointReport()

report.add_title_slide(title="An automatically generated presentation")
report.add_slide("content/lion.jpg", title="One image")
report.add_slide(["content/lion.jpg", "Text related to the image"], title="Images and text")
report.add_slide("content/*.jpg", title="Multiple images")

report.write_config("example1.json")
report.save("example1.pptx", pdf=True)
