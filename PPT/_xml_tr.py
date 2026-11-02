# -*- coding: utf-8 -*-
import zipfile, re
files = {
 "NEW": r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版_自动解说.pptx",
 "OLD": r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_自动解说.pptx",
}
for k, p in files.items():
    z = zipfile.ZipFile(p)
    names = [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)]
    s1 = z.read("ppt/slides/slide1.xml").decode("utf-8")
    has_tr = "<p:transition" in s1
    m = re.search(r"<p:transition[^>]*>", s1)
    print(k, "slides:", len(names), "| transition:", has_tr, "| tag:", m.group(0) if m else None)
    print("   tail order:", re.findall(r"<p:(cSld|clrMapOvr|transition|timing|extLst)", s1)[:8])
    z.close()