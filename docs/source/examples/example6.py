
from pptreport import PowerPointReport

report = PowerPointReport()

report.add_title_slide(title="Different margins")
report.add_slide(content=["content/*.jpg"],
                 inner_margin=0,
                 title="A grid (no inner margins)", n_columns=3)
report.add_slide(content=["content/*.jpg"],
                 outer_margin=0,
                 title="A grid (no outer margins)", n_columns=3)
report.add_slide(content=["content/zebra_fish.png"],
                 left_margin=0,
                 right_margin=4,
                 title="Custom margins")

report.write_config("example6.json")
report.save("example6.pptx", pdf=True)
