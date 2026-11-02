from pptx import Presentation
import os
src = r"D:\doc\ZGCLab\triza\project\TRIZa_答辩PPT_1018.pptx"
outdir = r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022\assets"
prs = Presentation(src)
n = 0
for idx in (25, 26):
    s = prs.slides[idx-1]
    for sh in s.shapes:
        if sh.shape_type == 13:
            img = sh.image
            ext = img.ext
            n += 1
            fn = os.path.join(outdir, f"principle_{idx}_{n}.{ext}")
            with open(fn, "wb") as f:
                f.write(img.blob)
            print(fn, img.size, img.dpi, round(sh.width/914400,2), round(sh.height/914400,2))
