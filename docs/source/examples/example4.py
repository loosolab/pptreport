from pptreport import PowerPointReport

report = PowerPointReport()

report.add_title_slide(title="Custom content layout")
content_layout = [[0, 1, 2],
                  [3, 3, 3]]
report.add_slide(content_layout=content_layout,
                 title="Custom content layout 1",
                 content=["content/mandarin_fish.jpg", "content/clown_fish.jpg",
                          "content/blue_tang_fish.jpg", "content/zebra_fish.png"])
content_layout = [[0, 2, 3],
                  [1, 2, 4]]
report.add_slide(content_layout=content_layout,
                 title="Custom content layout 2",
                 content=["content/mandarin_fish.jpg", "content/clown_fish.jpg", "content/giraffe.jpg",
                          "content/blue_tang_fish.jpg", "content/zebra_fish.png"])
content_layout = [[0, 3],
                  [1, 3],
                  [2, -1]]  # use -1 to keep position empty
report.add_slide(content_layout=content_layout,
                 title="Custom content layout 3",
                 content=["content/mandarin_fish.jpg", "content/clown_fish.jpg", "content/blue_tang_fish.jpg",
                          "content/giraffe.jpg"])

report.write_config("example4.json")
report.save("example4.pptx", pdf=True)
