Download PyMuPDF wheel for installed python version and platform from
    https://pypi.org/project/PyMuPDF/#files
and install like following
    pip install PyMuPDF-1.18.0-cp38-cp38-win_amd64.whl

install other requirements via
    pip install -r requirements.txt

python main.py --folder ".\\PDF Working Data\\Sample 1" --logo logo.png --title "Complete Subscription Package" \
    --sub_title "As Of October 9, 2020" --comment "This is a paragraph which should appear after the sub-title,
    but before the start of the links/items in the table of contents"

By default, it will process files in Sample 2 directory and and produce merged pdf file named 'Sample 2.pdf'(directory
name) with above parameters.
