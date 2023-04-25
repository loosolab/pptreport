from pptreport import PowerPointReport

report = PowerPointReport()

report.add_title_slide(title="Change fontsize")
content_layout = [[0, 2],
                  [1, 2]]
report.add_slide(content=["content/fish_description.md"] * 3,
                 width_ratios=[0.2, 0.8],
                 title="Automatic fontsize",
                 content_layout=content_layout)
report.add_slide(content=["content/fish_description.md"] * 3,
                 title="Fontsize set to 11.5",
                 fontsize=11.5,
                 content_layout=content_layout)

report.write_config("example14.json")
report.save("example14.pptx", pdf=True)
