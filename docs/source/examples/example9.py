from pptreport import PowerPointReport

report = PowerPointReport(template="content/template.pptx")

report.add_title_slide(title="Advanced RegEx usage")
report.add_slide(title="Globbing with *", content="*/*.png")
report.add_slide(title="Regex character matching with .*", content="content/.*.png")
report.add_slide(title="Regex character match", content="content/[a-c]+.*.jpg", show_filename=True)
# files named either zebra_fish or horse_fish with either jpg or png ending
report.add_slide(title="Regex OR (zebra|horse)", content="content/(zebra|horse)_fish.(png|jpg) ")
# regex in directories
report.add_slide(title="Regex .*", content="content/.*_animals/.*.jpg")
# regex in directories
report.add_slide(title="Regex '.*' matches upper directories too", content=".*_animals/.*.jpg")

report.write_config("example9.json")
report.save("example9.pptx", pdf=True)
