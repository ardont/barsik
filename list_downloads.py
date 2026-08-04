import glob
import os

downloads_pattern = r"C:\Users\Maxim\Downloads\*13-19*"
print("Searching for pattern:", downloads_pattern)
for f in glob.glob(downloads_pattern):
    print("Found file repr:", repr(f))
