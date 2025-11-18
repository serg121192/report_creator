import argparse

parser = argparse.ArgumentParser(description="Choose the action:\n")
parser.add_argument("--user-report", action="store_true")
parser.add_argument("--block-inactive", action="store_true")

args = parser.parse_args()