#!.venv/bin/python3
#-*- coding: UTF-8 -*-

__author__ = 'helljump'

#  Title Category Permalink Meta Description Content Date

import codecs
import re
import xlwt
from datetime import datetime
from dateutil.parser import parse
import argparse
from os.path import splitext


parser = argparse.ArgumentParser()
parser.add_argument("finp", help="wp xml file")
args = parser.parse_args()


def parse_xml(data):
    for item in re.finditer(r"(?imsu)<item>(.+?)</item>", data):
        egg = {'text': u''}
        if (not re.search("(?imsu)<wp:post_type>post</wp:post_type>", item.group(1)) and
                not re.search("(?imsu)<wp:post_type>page</wp:post_type>", item.group(1))):
            continue

        m = re.search("(?imsu)<link>(.+?)</link>", item.group(1))
        if m:
            egg["link"] = m.group(1).split('/')[-1]

        m = re.search("(?imsu)<description>(.+?)</description>", item.group(1))
        egg["description"] = m.group(1) if m else ''
        m = re.search("(?imsu)<title>(.+?)</title>", item.group(1))
        if m:
            egg["title"] = m.group(1)
        m = re.search("(?imsu)<wp:post_date>(.+?)</wp:post_date>", item.group(1))
        if m:
            egg["date"] = parse(m.group(1))
        m = re.search("(?imsu)<content:encoded><!\[CDATA\[(.+?)\]\]></content:encoded>", item.group(1))
        if m:
            egg["text"] = m.group(1).replace('<!--more-->', '<hr class="more">')
        m = re.search("(?imsu)<excerpt:encoded><!\[CDATA\[(.+?)\]\]></excerpt:encoded>", item.group(1))
        if m and len(m.group(1)) > 3:
            egg['text'] = egg['text'].replace('<hr class="more">', '')
            egg["text"] = "%s\n<hr class='more'>\n%s" % (m.group(1), egg['text'])
        m = re.search("(?imsu)<category.+?><!\[CDATA\[(.+?)\]\]></category>", item.group(1))
        if m:
            egg["category"] = m.group(1)
        else:
            egg["category"] = u'Новости'
        m = re.search("(?imsu)<wp:post_id>(\d+)</wp:post_id>", item.group(1))
        if m:
            egg["post_id"] = m.group(1)
        m = re.search("(?imsu)<wp:post_parent>(\d+)</wp:post_parent>", item.group(1))
        if m:
            egg["post_parent"] = m.group(1)
        m = re.findall("(?imsu)<category domain=\"(post_tag|tag)\".+?><!\[CDATA\[(.+?)\]\]></category>", item.group(1))
        if m:
            egg["tags"] = [row[1] for row in m]
        else:
            egg["tags"] = []
        yield(egg)


date_style = xlwt.XFStyle()
date_style.num_format_str = 'DD-MM-YY'

wb = xlwt.Workbook()
ws = wb.add_sheet(u'Page 1')

finp = codecs.open(args.finp, "r", "utf-8").read()

for i, t in enumerate("Title,Category,Permalink,MetaDescription,Content,Tags,Assigned Keywords,Date".split(',')):
    ws.write(0, i, t)

for row, data in enumerate(parse_xml(finp), 1):
    ws.write(row, 0, data['title'])
    ws.write(row, 1, data['category'])
    ws.write(row, 2, data['link'])
    ws.write(row, 3, '')
    ws.write(row, 4, data['text'])
    ws.write(row, 5, ', '.join(data['tags']))
    ws.write(row, 6, '')
    ws.write(row, 7, data['date'], date_style)

fn, fe = splitext(args.finp)
wb.save(fn + '.xls')
