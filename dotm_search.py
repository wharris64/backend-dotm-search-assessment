#!/usr/bin/env python
# -*- coding: utf-8 -*-
import argparse
import os
import re
from zipfile import ZipFile
"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Will Harris"


def initparser():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dir", default=os.getcwd())
    parser.add_argument("charsearch")

    return parser


def search(dir, char):
    print(os.getcwd())
    count = 0
    file_count = 0
    matches = {}
    os.chdir(dir)
    for filename in os.listdir(os.getcwd()):
        file_count += 1
        if filename.endswith(".dotm"):
            with ZipFile(filename).open("word/document.xml") as doc:
                for line in doc:
                    
                    looksee = re.search("((.{1,20}|)\\"+char+"(.{1,20}|))", line)

                    try:
                        matches[filename] = looksee.group()
                        count += 1
                        continue
                    except:
                        pass
    return [matches, count, file_count]

def main():
    parser = initparser()
    args = parser.parse_args()
    result, count, file_count = search(args.dir, args.charsearch)
    for key in list(result.keys()):
        print(key, result[key])
    print(str(count), str(file_count))
    
        

    # print(parser.parse_args())
    #raise NotImplementedError("Your awesome code begins here!")


if __name__ == '__main__':
    main()
