# MSExcelArt
Drawing with Microsoft Excel

## Usage
### Package requirement
Please install all required packages first if you do not have `opencv` and `openpyxl`.
```
pip install -r requirements.txt
```
## Draw a 2D image
```
python3 main_image.py --image $image_source
```
With this command, a `output.xlsx` will be generated.

## Generate a video
```
python3 main_video.py --video $video_source
```
With this command, a `output.xlsx` will be generated.