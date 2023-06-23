from pptreport import PowerPointReport

report = PowerPointReport(template="content/template.pptx")
report.add_global_parameters({"content_alignment": "left"})

report.add_title_slide(title="Examples of markdown usage")
report.add_slide(content=["content/*_fish*", "content/fish_description.md"],
                 title="Fish with text from markdown")
report.add_slide(content=["content/blue_tang_fish.jpg", "This is a __fish__"],
                 title="Fish with markdown formatted string")

report.add_slide(content=["content/headers.md"], title="Headers")
report.add_slide(content=["content/text.md"], title="Text formatting")
report.add_slide(content=["content/lists.md"], title="Lists")

report.write_config("example13.json")
report.save("example13.pptx", pdf=True)
