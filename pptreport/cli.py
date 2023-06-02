import sys
import argparse
from pptreport import __version__ as pptreport_version
from pptreport import PowerPointReport


def main():

    # Parse args
    parser = argparse.ArgumentParser("pptreport")
    parser.add_argument("--config", metavar="<path>", help="Path to configuration file in json format, e.g. config.json", required=True)
    parser.add_argument("--output", metavar="<path>", help="Path to output file, e.g. presentation.pptx", required=True)
    parser.add_argument("--template", metavar="<path>", help="Path to template ppt file (optional). Will overwrite any template specified in config file.")
    parser.add_argument("--pdf", help="Additionally save the presentation as a .pdf with the same basename as --output. Requires 'Libreoffice' to be installed on path (Default: False).", action='store_true', default=False)
    parser.add_argument("--show-borders", help="Show borders around all elements in the presentation. Good for debugging layouts (Default: False).", action='store_true', default=False)
    parser.add_argument("--verbosity", metavar="0/1/2", help="Verbosity level for logging (0-2). 0 = only errors and warnings, 1 = minimal logging, 2 = debug logging (Default: 1).", type=int, default=1)
    parser.add_argument("--version", action="version", version=pptreport_version)

    # If no args, print help
    if len(sys.argv[1:]) == 0:
        parser.print_help()
        sys.exit()

    args = parser.parse_args()

    # Read config file
    import json
    with open(args.config) as f:
        try:
            config_dict = json.load(f)
        except Exception as e:
            print(f"Error reading config file '{args.config}'. Error was: {e}")
            sys.exit(1)

    # Overwrite template if specified
    if args.template:
        config_dict["template"] = args.template

    # If global show_borders was given, overwrite config
    if args.show_borders:
        if "global_parameters" not in config_dict:
            config_dict["global_parameters"] = {}
        config_dict["global_parameters"]["show_borders"] = args.show_borders

    # Create report using PowerPointReport class
    report = PowerPointReport(verbosity=args.verbosity)
    try:
        report.from_config(config_dict)
        report.save(args.output, pdf=args.pdf)

    except Exception as e:
        report.logger.error(e)  # show exception
        sys.exit(1)

    report.logger.info("pptreport finished!")
