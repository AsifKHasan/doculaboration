from PIL import Image
from pprint import pprint

# photo_path = '/home/asif/projects/asif@github/doculaboration/out/tmp/photo__Ahmed.Nafis.Fuad.png'
photo_path = '/home/asif/projects/asif@github/doculaboration/out/tmp/photo__Md.Jakir.Hossain.png'

im = Image.open(photo_path)

info = im.info
dpi = im.info['dpi']

type(dpi[0])
