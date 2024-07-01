Image compression using Pillowscript

Prerequisites
-------------

Programmatic compression of images using existing algorithmscan be undertaken using a command-line script.

1. Download the [lossy\_compression.py](https://drive.google.com/file/d/1o_VrAd2Dp6Uf_fRSCycsiFQej2jF50qV/view?usp=drive_link) script to your device
2. Move the script to a stable location, such as your userhome directory
3. Make the script executable `chmod+x .../lossy_compression.py`
 - Install [ImageMagick](http://www.imagemagick.org/script/index.php) `brew install imagemagick`
 - Install [Pillow](https://pillow.readthedocs.io/en/stable/) `python3 -m pip install pillow`
 - Install [Numba](https://numba.pydata.org/) `python3 -mpip install numba`
 - Install [NumPy](https://numpy.org/) `python3 -mpip install numpy`
 - Install [openpyxl](https://openpyxl.readthedocs.io/en/stable/) `python3 -m pip install openpyxl`

1\. Bulk resize
---------------

To bulkresize all images in a directory, use Mac’s built-in [sips](https://ss64.com/osx/sips.html)tool. Navigate to the destination in Terminal and then run the followingscript:

`sips -Z 1200 \*.jpg`

\*NB this script may inadvertently upscale images smallerthan the dimensions specified.

To avoid this, navigate to a top-level directory usingTerminal and use the following to search for all files that exceed thespecified dimension recursively, and then only apply the transformation tothose:

`mogrify -resize '2000>' *.jpg `

2\. Bulk set DPI
----------------

To bulk change the DPI of all images in a directory,navigate to the destination in Terminal and then run the following script:

`find . -name "*.jpg" -exec mogrify -units "PixelsPerInch" -density 72\> {} \;`

3\. Dithering
-------------

Output formats:
- JPEG
- PNG
- WebP

Dithering modes:
- one\_bit
- bayer\_2x2
- cluster\_4x4
- bayer\_8x8
- floyd\_steinberg\_pil
- atkinson
- stucki
- jarvis\_judice\_ninke
- grayscale
- floyd\_stinberg\_dev

`python3 "SCRIPT LOCATION" "IMAGE DIRECTORY" "lossy_compression" "DITHERING MODE" "FORMAT"`

The script will create a new directory at the same level asthe source in which to store the derivatives generated. Additionally, it willproduce a spreadsheet listing the files and their relative transformation sizes.
