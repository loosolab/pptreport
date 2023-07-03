from pptreport import PowerPointReport

report = PowerPointReport(template="content/template.pptx")

report.add_title_slide(title="Examples of grouped content")

report.add_slide(title="Three images per group", slide_layout="Section Header")
report.add_slide(grouped_content=["content/colored_animals/(\w+)_blue.jpg",
                                  "content/colored_animals/(\w+)_yellow.jpg",
                                  "content/colored_animals/(\w+)_red.jpg"],
                 n_columns=3, missing_file="empty")

report.add_slide(title="Three images per group + show_filename", slide_layout="Section Header")
report.add_slide(grouped_content=["content/colored_animals/(\w+)_blue.jpg",
                                  "content/colored_animals/(.*)_yellow.jpg",
                                  "content/colored_animals/([a-z\_]+)_red.jpg"],
                 n_columns=3, missing_file="empty", show_filename=True)

report.add_slide(title="Three images per group (missing_file='text')", slide_layout="Section Header")
report.add_slide(grouped_content=["content/colored_animals/(\w+)_blue.jpg",
                                  "content/colored_animals/(.*)_yellow.jpg",
                                  "content/colored_animals/([a-z\_]+)_red.jpg"],
                 n_columns=3, missing_file="text")

report.add_slide(title="Two images per group + strings", slide_layout="Section Header")
report.add_slide(grouped_content=["Blue animal", "Yellow animal",
                                  "content/colored_animals/(\w+)_blue.jpg",
                                  "content/colored_animals/(\w+)_yellow.jpg"],
                 missing_file="empty")

report.write_config("example10.json")
report.save("example10.pptx", pdf=True)
