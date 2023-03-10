from importlib import import_module

from ._version import __version__

# Set functions to be available directly
global_classes = ["pptreport.powerpointreport.PowerPointReport"]

for c in global_classes:

    module_name = ".".join(c.split(".")[:-1])
    attribute_name = c.split(".")[-1]

    module = import_module(module_name)
    attribute = getattr(module, attribute_name)

    globals()[attribute_name] = attribute
