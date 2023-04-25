from pptreport import PowerPointReport

report = PowerPointReport()

report.add_title_slide(title="vertical/horizontal content layout")
report.add_slide(content=["content/lion.jpg", "Some text below the picture."],
                 content_layout="vertical",
                 title="A lion (vertical layout)")
report.add_slide(content=["content/lion.jpg", "Using 'height_ratios' controls how much vertical space the picture has."],
                 content_layout="vertical",
                 height_ratios=[0.9, 0.1],
                 title="A lion (specific height ratios)")
report.add_slide(content=["content/lion.jpg", "Some text next to the picture."],
                 content_layout="horizontal",
                 title="A lion (horizontal layout)")
report.add_slide(content=["content/lion.jpg", "Using 'width_ratios' controls how much vertical space the picture has."],
                 content_layout="horizontal",
                 width_ratios=[0.8, 0.2],
                 title="A lion (specific width ratios)")

report.write_config("example3.json")
report.save("example3.pptx", pdf=True)
