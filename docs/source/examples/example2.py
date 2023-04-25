from pptreport import PowerPointReport

report = PowerPointReport()


report.add_title_slide(title="Control images per slide")
report.add_slide(content="content/*_fish*",
                 title="With split=True, each fish gets their own slide",
                 split=True)
report.add_slide(content=["content/colored_animals/dog*", "content/colored_animals/lion*", "content/colored_animals/mouse*"],
                 title="With split=3, every slide contains 3 pictures",
                 n_columns=3,
                 split=3)

report.write_config("example2.json")
report.save("example2.pptx", pdf=True)
