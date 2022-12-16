import sys
import argparse
from pptreport import __version__ as pptreport_version
from pptreport.classes import PowerPointReport

def main():

    # Parse args
    parser = argparse.ArgumentParser("pptreport")
    parser.add_argument("--config", help="Path to config file in json format, e.g. config.json", required=True)
    parser.add_argument("--output", help="Path to output file, e.g. presentation.pptx", required=True)
    parser.add_argument("--template", help="Path to template ppt file (optional). Will overwrite any template specified in config file.")
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

    # Create report using PowerPointReport class
    print("Building report...")
    report = PowerPointReport()
    report.from_config(config_dict)
    report.save(args.output)

    print(f"Report saved to {args.output}")

if __name__ == "__main__":
    main()
