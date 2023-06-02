from pptreport import PowerPointReport

report = PowerPointReport(template="content/template.pptx")
report.add_title_slide(title="Adjust 'max_pixels' for image content")

for pixels in ["1e3", "1e4", "1e5", "1e6"]:
    report.add_slide(content="content/clown_fish.jpg", max_pixels=pixels, title=f"max_pixels={pixels}")

report.write_config("example17.json")
report.save("example17.pptx", pdf=True)
