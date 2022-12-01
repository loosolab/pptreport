import sys
import argparse
from pptreport import __version__ as pptreport_version


def main():

    print("Hello World!")
    print(pptreport_version)

    parser = argparse.ArgumentParser("pptreport")

    parser.add_argument("--config", help="Path to config file", required=True)
    parser.add_argument("--output", help="Path to output file", required=True)
    parser.add_argument("--template", help="Path to template ppt file")
    parser.add_argument("--version", action="version", version=pptreport_version)

    # If no args, print help
    if len(sys.argv[1:]) == 0:
        parser.print_help()
        sys.exit()


def build_report(args):
    pass
