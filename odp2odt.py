#!/usr/bin/python3

import sys, os
import zipfile
from xml.dom.minidom import parse, parseString

def get_text(node):
    if node.nodeType == node.TEXT_NODE: return (node.nodeValue or '').strip()
    elif node.nodeType == node.ELEMENT_NODE: return ' '.join(filter(None, map(get_text, node.childNodes)))
    else: return ''

def convert(inp, out):
    in_zip = zipfile.ZipFile(inp, 'r')
    in_zip.testzip()
    empty_zip = zipfile.ZipFile(os.path.join(os.path.dirname(__file__), 'empty.odt'), 'r')
    empty_zip.testzip()
    in_dom = parseString(in_zip.read('content.xml').decode('utf-8'))
    out_dom = parseString(empty_zip.read('content.xml').decode('utf-8'))
    out_text = out_dom.getElementsByTagName('office:text')[0]
    last_hdr = None
    for page in in_dom.getElementsByTagName('draw:page'):
        for box in page.getElementsByTagName('draw:text-box'):
            head_nest = ('list', 'list-header', 'p')
            cur = box
            hdr = None
            for tag in head_nest:
                if len(cur.childNodes) == 1 and cur.firstChild.localName == tag:
                    cur = cur.firstChild
                    if tag == 'list' and cur.getAttribute('text:style-name') != 'L1': break
                else:
                    break
            else:
                hdr = get_text(cur).replace('\n', ' ')
            if hdr:
                if hdr != last_hdr:
                    print("Found heading: ", hdr)
                    h = out_dom.createElement('text:h')
                    h.setAttribute('text:style-name', "Heading_20_1")
                    h.setAttribute('text:outline-level', "1")
                    h.appendChild(out_dom.createTextNode(hdr))
                    out_text.appendChild(h)
                    last_hdr = hdr
                else:
                    print('Ignoring repeated heading', hdr)
            else:
                for ch in box.childNodes:
                    out_text.appendChild(ch)
        for frame in page.getElementsByTagName('draw:frame'):
            if frame.firstChild and frame.firstChild.tagName == 'draw:image':
                frame.removeAttribute('svg:x')
                frame.removeAttribute('svg:y')
                frame.removeAttribute('draw:layer')
                frame.removeAttribute('draw:style-name')
                frame.removeAttribute('draw:text-style-name')
                frame.setAttribute('text:anchor-type', "as-char")
                p = out_dom.createElement('text:p')
                p.appendChild(frame)
                out_text.appendChild(p)
    out_zip = zipfile.ZipFile(out, 'w')
    out_zip.writestr('content.xml', out_dom.toxml('utf-8'))
    for thezip in (empty_zip, in_zip):
        for info in thezip.infolist():
            try: out_zip.getinfo(info.filename)
            except KeyError: out_zip.writestr(info.filename, thezip.read(info))
    out_zip.close()

if __name__ == '__main__':
    inp = sys.argv[1]
    name, ext = os.path.splitext(inp)
    if ext != '.odp': raise RuntimeError("Only ODP files supported")
    out = name + '.odt'
    convert(inp, out)

