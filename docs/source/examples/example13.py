from pptreport import PowerPointReport

report = PowerPointReport()

report.add_title_slide(title="Example of markdown usage")
report.add_slide(content=["content/*_fish*", "content/fish_description.md"],
                 title="Fish with text from markdown")
report.add_slide(content=["content/blue_tang_fish.jpg", "This is a __fish__"],
                 title="Fish with markdown formatted string")

report.write_config("example13.json")
report.save("example13.pptx", pdf=True)
