from pptreport import PowerPointReport

report = PowerPointReport(template="content/template.pptx")

report.add_title_slide(title="Filenames above images")
report.add_slide(content="content/numbered_animals/giraffe*.jpg",
                 title="Filenames per image",
                 n_columns=3,
                 show_filename=True)
report.add_slide(content="content/numbered_animals/giraffe*.jpg",
                 title="Alignment of filenames per image",
                 n_columns=3,
                 show_filename=True,
                 filename_alignment="left")
report.add_slide(content="content/numbered_animals/giraffe*.jpg",
                 title="With show_filename=filepath",
                 n_columns=3,
                 show_filename="filepath")

report.write_config("example8.json")
report.save("example8.pptx", pdf=True)
