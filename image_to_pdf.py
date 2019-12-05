import img2pdf

#imagelist is the list with all the image filenames
a4inpt = (img2pdf.mm_to_pt(210),img2pdf.mm_to_pt(297))
layout_fun = img2pdf.get_layout_fun(a4inpt)
imglist = ["test-1.png","test-2.png"]
def convert(imglist):
    with open('temp.pdf','wb') as f:
        f.write(img2pdf.convert(imglist,layout_fun=layout_fun))

if __name__ == "__main__":
    convert(imglist)
