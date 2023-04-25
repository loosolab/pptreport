from pptreport import PowerPointReport

report = PowerPointReport(template="content/template.pptx")

report.add_title_slide(title="Control content alignment")
report.add_slide(content=["content/zebra_fish.png"] * 3,  # three times the same picture
                 content_alignment=["lower", "center", "upper"],
                 title="Different alignments on the same slide",
                 n_columns=3)

for horizontal_alignment in ["left", "center", "right"]:
    report.add_slide(content=["content/giraffe.jpg"],
                     content_alignment=horizontal_alignment,
                     title=f"Example of {horizontal_alignment} alignment")

# Alignment of text
content_alignments = ["upper left", "upper center", "upper right",
                      "center left", "center", "center right",
                      "lower left", "lower center", "lower right",
                      ]
texts = [f"'{align}' alignment" for align in content_alignments]

report.add_slide(content=texts,
                 content_alignment=content_alignments,
                 n_columns=3,
                 title="Text alignments")


report.write_config("example5.json")
report.save("example5.pptx", pdf=True)
