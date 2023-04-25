from pptreport import PowerPointReport

report = PowerPointReport()

report.add_title_slide(title="Order of content")
report.add_slide(content=["First giraffe", "Second giraffe", "Third giraffe",
                          "content/numbered_animals/giraffe*.jpg"],
                 n_columns=3,
                 height_ratios=[0.1, 0.9],
                 title="Files are expanded with natural sorting")
report.add_slide(content=["First giraffe", "Second giraffe", "Third giraffe",
                          "content/numbered_animals/giraffe*.jpg"],
                 title="The option 'fill_by' controls the fill order",
                 fill_by="column")

report.write_config("example7.json")
report.save("example7.pptx", pdf=True)
