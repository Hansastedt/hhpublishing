#!/usr/bin/env python3

import os
import argparse
import hashlib
import base64 
import datetime


try:
    import mammoth
except:
    print("Please install required library: mammoth")
    print("Try: pip3 install mammoth")
    exit(1)

parser = argparse.ArgumentParser()

parser.add_argument("sourcedir", help="Source directory.")
parser.add_argument("outputdir", help="Output directory.")

args = parser.parse_args()



if not os.path.isdir(args.sourcedir):
    print("Source dir is not a directory.")
    exit(1)

if not os.path.isdir(args.outputdir):
    print("Output dir is not a directory.")
    exit(1)



def hashfile(path):
    d = hashlib.sha256(open(path, "rb").read()).digest()
    return base64.b32encode(d).decode("ascii").strip("=").lower()

def get_output_filename(path):
    h = hashfile(path)
    ct = datetime.datetime.fromtimestamp(os.path.getctime(path))
    ct = ct.strftime("%Y-%m-%d")
    return "%s-%s.html" % (ct, h)


def convert(path):
    with open(path, "rb") as docx:
        result = mammoth.convert_to_html(docx)
        return result.value

def decorate_html(text):
    return ("""
---
title: test
layout: default
---
%s""" % text).strip()


sourcelist = [
    e for e in os.listdir(args.sourcedir)
    if e.lower().endswith(".docx")
]

for sourcefile in sourcelist:
    sourcepath = os.path.join(args.sourcedir, sourcefile)
    outputfilename = get_output_filename(sourcepath)
    outputfullpath = os.path.join(args.outputdir, outputfilename)

    open(outputfullpath, "w+").write(decorate_html(convert(sourcepath)))
