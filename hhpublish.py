#!/usr/bin/env python3

helpinfo = """
Tool for easy publishing articles from a source directory to Jekyll based
website.

This tool converts *.docx files within source directory to Jekyll conformed
posts in a target directory.

Usage:

    ./hhpublish.py --help   Check for usage information.
    ./hhpublish.py <sourcedir> <outputdir>
                            Converts any *.docx file in <sourcedir> to a
                            corresponding file in <outputdir>.

During conversion:

1. Embedded images of docx will be stored inline as Data URLs in HTML.
2. Only modified source file will be converted. Any file in target directory
   that does not match a source file will be deleted. This is done by
   corresponding source files' hash digests with target filenames.

"""

import os
import argparse
import hashlib
import base64 
import datetime


try:
    import mammoth
    import bs4
    from bs4 import BeautifulSoup
    import yaml
except:
    print("Try run first with:\n  pip3 install beautifulsoup4 PyYAML mammoth")
    exit(1)










parser = argparse.ArgumentParser(description=helpinfo)

parser.add_argument("sourcedir", help="Source directory.")
parser.add_argument("outputdir", help="Output directory.")
parser.add_argument(
    "--dry-run", action="store_true", help="Test run without modifying files.")

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
    return h, ("%s-%s.html" % (ct, h))


def convert(path):
    with open(path, "rb") as docx:
        result = mammoth.convert_to_html(docx)
        return result.value

##############################################################################

def decorate_html(text):
    metadata = {
        "layout": "default",
        "categories": [],
    }
    htmldoc = BeautifulSoup(text, 'html.parser')

    try:
        tables = htmldoc.find_all('table')
        first_table = tables[0]
    except:
        raise Exception("没有找到元数据定义表格。")

    try:
        rawdata = [
            tr.find_all('td')[:2]
            for tr in first_table.find_all('tr')
        ]
        rawdata = [
            (
                "\n".join(e[0].strings).strip().lower(),
                "\n".join(e[1].strings).strip()
            )
            for e in rawdata
        ]
        rawdata = [ e for e in rawdata if e[0] or e[1] ]
    except Exception as e:
        print(e)
        raise Exception("元数据表格定义错误。")
    
    for key, value in rawdata:
        if key in ['标题', 'title', 'titel']:
            metadata['title'] = value
        elif key in ['作者', 'author', 'autor']:
            metadata['author'] = value
        elif key in ['分类', 'category', 'categories']:
            metadata['categories'].append(value.lower().strip())

    metadata["categories"] = list(set(metadata["categories"]))

    first_table.decompose()
    return ("""
---
%s
---
%s""" % (yaml.dump(metadata, allow_unicode=True), str(htmldoc))).strip()

##############################################################################

sourcelist = [
    e for e in os.listdir(args.sourcedir)
    if e.lower().endswith(".docx") and not e.startswith("~")
]

targetlist = [
    e
    for e in os.listdir(args.outputdir)
    if e.lower().endswith(".html")
]

targetdeletelist = [e for e in targetlist]

def hash_is_in_target(h):
    for t in targetlist:
        if h in t: return t 
    return False


for sourcefile in sourcelist:
    sourcepath = os.path.join(args.sourcedir, sourcefile)
    filehash, outputfilename = get_output_filename(sourcepath)

    target_matching_filename = hash_is_in_target(filehash)
    if False != target_matching_filename:
        targetdeletelist.remove(target_matching_filename)
        print("Ignoring: %s" % sourcefile)
        continue # exists, no process for this file
    else:
        outputfullpath = os.path.join(args.outputdir, outputfilename)
        try:
            converted_html = convert(sourcepath)
            decorated_html = decorate_html(converted_html)
            if not args.dry_run:
                open(outputfullpath, "w+").write(decorated_html)
            print("*******\n\n" + decorated_html)
            print("Generated for: %s" % sourcefile)
        except Exception as e:
            print(e)
            pass # TODO here write the error into output

for tobedeleted in targetdeletelist:
    deleted_filepath = os.path.join(args.outputdir, tobedeleted)
    if not args.dry_run:
        os.remove(deleted_filepath)
    print("Delete: %s" % deleted_filepath)
