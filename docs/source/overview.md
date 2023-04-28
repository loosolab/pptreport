# pptreport - automatic creation of powerpoint presentations

The pptreport package is a tool for building powerpoint presentations using a configuration file of content such as pictures and text, or step-by-step during a script or jupyter notebook. 

## How to install

pptreport can be installed from github using:
```bash
$ git clone https://gitlab.gwdg.de/loosolab/software/pptreport.git
$ cd pptreport 
$ pip install .
```

## Commandline usage
An example of command-line usage is:
```bash
$ cd examples/
$ pptreport --config report_config.json --output report.pptx
```


## Python usage
Examples of usage within python are found in the section [Examples](examples/index.rst).

In summary, a report is initialized, slides are added and the presentation is finally saved as seen here:
```python
#Initialize presentation
report = PowerPointReport()

#Add slides
report.add_title_slide(title="An automatically generated presentation")
report.add_slide("lion.jpg", title="One image")
report.add_slide(["lion.jpg", "Text related to the image"], title="Images and text")
report.add_slide("*.jpg", title="Multiple images")

#Save presentation
report.save("report.pptx")
```


## Configuration file manual

A report can be built using a json-formatted configuration file with the format:
```
{ template: "template.pptx,
  slides: [
            {   <configuration for slide 1>   },
            {   <configuration for slide 2>   },
            {   <configuration for slide 3>   }
          ]
}
```
Where the the configurations have the format:
```
{ title: "Title of the slide", content: ["image.png", "another_image.png"] }
```
The example file at [examples/report_config.json](../../examples/report_config.json) gives various examples of usage. For further reference, the tables below give more detailed descriptions of valid keys for each level.

<br />
 
### Configuration of the presentation:

| Key | Type | Description | Example |
| --- | ---- | -------- | ------ |
| template |  string | This is the path to an optional template to use for the presentation, for example to use a specific slide design. The presentation is initialized with the slides of the presentation. In order to use the slide master exclusively, delete all slides from the presentation. | "mytemplate.pptx" |
| size | string or list of length 2| If template is not given, size controls the size of the presentation. Can be "standard" (default), "widescreen", "a4-portait", "a4-landscape". Can also be a list of two numbers indicating [height, width] in cm, e.g. [21, 14.8] for A5 size. | "standard"
| slides | list of dictionaries | This key contains a list of configuration dictionaries. Each dictionary corresponds to one slide in the presentation. | See configuration of slide keys below. |  
| global_parameters | dictionary | A dictionary containing slide keys (see below), which will be global across all slides. These will be overwritten by any specific configuration given per slide. | {"inner_margin": 0} |

<br />
<br />

### Configuration of slides:

| Key | Type | Description | Examples |
| --- | ---- | ------- | ------ |
| content | list of strings | List of content to be added to the slide. These can be a paths to files (either image or text) or strings to directly write to a text box. If the string contains "*", files matching the pattern will be found. | ["imagefile.png"]<br />["image\*.png"]<br />["Some text", "Additional text"]<br />["file.txt", "image.jpg"] |
| grouped_content | list of regex patterns | A list of regex patterns with groups. This option can be used inplace of 'content' to create slides with images per regex group. A regex group is given with parenthesis, e.g. if we have a list of heatmaps and scatterplots for a number of samples, we can create per-sample grouped slides using: grouped_content = ["heatmap_sample([0-9]+).png", "scatter_sample([0-9]+).png"]. If no "title" is given, the default slide title is "Group: \<regex-group-name\>". Note: the option "split=True" is not valid when grouped_content is given. | ["image1_(.+).png", "image2_(.+).png"] |
| title | str | The text to put in the placeholder for "title" on the slide. | "My title" |
| slide_layout | int or string| If an integer, this represents which master slide to be used for the layout. The numbering starts at 0, which is usually the title slide layout. The default is 1, which is usually a content slide with a title placeholder. However, note that this numbering depends on the template used. If 'slide_layout' is a string, the layout with that name is used. | 1<br />"Title Slide" |
| content_layout | string or list of ints | The layout of the content within the slide. Options are:<br />• "grid" (default): This places the content into a grid with `n_columns` columns.<br />• "vertical": Places content into one vertical column.<br />• "Horizonzal": Places contents into one horizontal row.<br /><br />The layout can also be specified as an array of integers, where each integer _i_ specified the _i_'th element in 'contents'. Numbering starts at 0. For example: <br />`[[0,0,0],`<br /> ` [1,1,2]]` <br /> describes a layout where the first element in contents is placed in the first row, the second is placed in the second row spanning two columns, and the third element in contents is placed in the lower right corner. Use '-1' to keep a position empty.  | "grid" |
| content_alignment | string or list of strings | The alignment of content within each content box. Can be vertical-horizontal combinations of "upper", "lower", "left", "right" and "center". If a string is given, all content on the slide will be aligned in the same way. If a list of strings is given, the alignment strings correspond to the order of elements in content, e.g. ["center", "left"] will align the first two elements center and left, respectively. The default is "center", which will align the content centered both vertically and horizontally. | "upper left"<br />"center"<br />"lower right"<br />["center", "left"] |
| n_columns | int | If "content_layout" is "grid", this integer controls how many columns to split the contents into. Default is 2. | 2 |
| outer_margin | float | The margin between slide border and the content. Default is 2 (cm). | 2<br />1.5 |
| inner_margin | float | The margin between the individual elements of contents. Default 1 (cm). | 1<br />0.5 |
| top_margin | float | Set the top outer margin. Default is the same as outer_margin. | 2 |
| bottom_margin | float | Set the bottom outer margin. Default is the same as outer_margin. | 2 |
| left_margin | float | Set the left outer margin. Default is the same as outer_margin. | 2 |
| right_margin | float | Set the right outer margin. Default is the same as outer_margin. | 2 |
| width_ratios | list of floats | This list of values, which must have the same length as the number of columns, controls the width of individual content columns. For example, `width_ratios=[2,1]` sets the first column to double the width of the second column. Default is that every column has equal width. | [0.1, 0.9]<br />[2, 2, 1]
| height_ratios | list of floats | This list of values, which must have the same length as the number of rows, controls the height of individual content rows. For example, `height_ratios=[2,1]` sets the first row to double the height of the second row. Default is that every row has equal height. | [0.1, 0.9]<br />[2, 2, 1] |
| notes | str | This string gives the path to a file, or the direct string, which should be added to the "notes" section of the slide. | "This text will be in notes"<br />"notes.txt" | 
| split | bool | If number of elements in `content` is large, it might be beneficial to split the elements to separate slides. If `split` is True, every element in content is written to its own slide (and as such, the slide config actually expands to more than one slide). All other parameters (such as title) are equal for all expanded slides. The default is False, meaning that all content is placed into one slide. | False |
| show_filename | bool or string | Show filenames above each image. The format of the displayed filename is decided by the options: <br />• True: filename without extension (e.g. 'cat') <br />• "filename": same as True <br />• "filename_ext": filename with extension (e.g. 'cat.jpg')<br />• "filepath": Full path to file (e.g. 'content/cat')<br />• "filepath_ext": full path to file with extension (e.g. 'content/cat.jpg')<br />• "path": the path to the file (e.g. 'content')<br />• False: no filename shown (default) | True<br />"filepath" |
| filename_alignment | string or list of strings | The horizontal alignment of the filename of an image within each content box. If a string is given, all filenames on the slide will be aligned in the same way. If a list of strings is given, the alignment strings correspond to the order of images in content, e.g. ["center", "left"] will align the first two filenames center and left, respectively. The default is "center", which will align the filename centered horizontally. | "left"<br />"center"<br />"right"<br />["center", "left"] |
| fill_by | str | The order of filling the content into the grid. Default is "row", in which case the content is filled row-by-row depending on 'n_columns' or the custom layout. The other option is "column", in which case the content is filled column-by-column. | "row" |
| remove_placeholders | bool | Whether to remove empty placeholders from the slide, e.g. if title is not given, powerpoint will show an empty text box. Default is False; to keep all placeholders. If True, empty placeholders will be removed. | True |
| fontsize | float | Fontsize of text content. If not given, the fontsize is automatically determined to fit the text in the textbox. | 12<br />10.5 |
| pdf_pages| int or "all" | Pages to include if pdf is a multipage pdf. "all" includes all available pages| "all"<br />[1, 3]<br />2|
|  missing_file | str | What to do if no files were found from a content pattern, e.g. "figure*.txt". Can be either "raise" (default), "text", "empty", "skip" or "skip-slide": <br />• "raise": a FileNotFoundError will be raised.<br />• "text": a content box will be added with the text of the missing content pattern<br />• "empty": an empty content box will be added for the missing content pattern.<br />• "skip": this content pattern will be skipped (no box added).<br />• "skip-slide": the whole slide will be skipped. | "raise"<br />"empty"<br />"skip" |

## Source of example images

All images and text are from Wikipedia and Wikimedia Commons:

- [blue_tang_fish.jpg](https://commons.wikimedia.org/wiki/File:Pacific_Blue_Tang_(Paracanthurus_hepatus)_(3149754704).jpg)
- [cat.jpg](https://commons.wikimedia.org/wiki/File:Cat_November_2010-1a.jpg)
- [dog.jpg](https://commons.wikimedia.org/wiki/File:Dog_(Canis_lupus_familiaris)_(1).jpg)
- [lion.jpg](https://commons.wikimedia.org/wiki/File:Lion_cubs_(51715160186).jpg)
- [mouse.jpg](https://commons.wikimedia.org/wiki/File:%D0%9C%D1%8B%D1%88%D1%8C_2.jpg)
- [giraffe.jpg](https://commons.wikimedia.org/wiki/File:Giraffe_2019-07-28.jpg)
- [clown_fish.jpg](https://commons.wikimedia.org/wiki/File:Amphiprion_ocellaris_(Clown_anemonefish)_by_Nick_Hobgood.jpg)
- [mandarin_fish.jpg](https://commons.wikimedia.org/wiki/File:Synchiropus_splendidus_2B_Luc_Viatour.jpg)
- [zebra_fish.png](https://commons.wikimedia.org/wiki/File:Danio_rerio_port.jpg)
- [chips.pdf](https://commons.wikimedia.org/wiki/File:Pommes-1.jpg)
- [fish_description.txt](https://en.wikipedia.org/wiki/Fish)
