import re

def extract_digits(s):
    return int(s) if s.isdigit() else s

def sort_key(filename):
    return [extract_digits(part) for part in re.split(r'(\d+)', filename)]

