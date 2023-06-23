import json
import os
import pptreport.cli
from unittest.mock import patch

content_dir = "examples/content"


def test_commandline():
    """ Test that the command line interface works as expected """

    config = {"template": "examples/content/template.pptx",
              "global_parameters": {"outer_margin": 1, "top_margin": 1.5},
              "slides": [{"title": "An automatically generated presentation", "slide_layout": 0},
                         {"title": "Layout can also be chosen using the layout name\n('Title Slide')",
                         "slide_layout": "Title Slide"},
                         {"content": f"{content_dir}/lion.jpg", "title": "A lion"},
                         {"content": [f"{content_dir}/dog.jpg", f"{content_dir}/cat.jpg"], "title": "Pets", "outer_margin": 3},
                         {"content": [f"{content_dir}/lion.jpg", "Some text below the picture."], "content_layout": "vertical", "title": "A lion (vertical layout)"}
                         ]}

    # Write to json file
    with open("test_config.json", "w") as f:
        json.dump(config, f)

    arguments = "pptreport --config test_config.json --output test_CLI.pptx".split(" ")
    with patch('sys.argv', arguments):
        pptreport.cli.main()

    assert os.path.exists("test_CLI.pptx")

    # Create report via API
    report = pptreport.PowerPointReport()
    report.from_config("test_config.json")
    report.save("test_API.pptx")

    # Assert that the two reports are the same
    assert os.stat("test_CLI.pptx").st_size == os.stat("test_API.pptx").st_size

    # Clean up
    os.remove("test_CLI.pptx")
    os.remove("test_API.pptx")
    os.remove("test_config.json")
