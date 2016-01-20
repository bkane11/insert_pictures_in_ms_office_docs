from __future__ import print_function

import add_images
import argparse

parser = argparse.ArgumentParser(prog='ADD IMAGES')
parser.add_argument('filepath', type=str, help='filepath for xlsx or docx file')
parser.add_argument('--images', nargs='+', help='list of filepaths for images to insert into xlsx file, separated by space', default=[])

args = parser.parse_args()

images = args.images
images = reduce(lambda a,b:a+b, map(lambda x: x.split(','), args.images) )

print('updating:', args.filepath)
print('adding:', args.images)

add_images.addImagesToMultipleSheets(name=args.filepath, images=images, then=lambda x: print({'success': x}))

# test command:
# python python/add_image_module.py tests/outputs/test.xlsx --images tests/images/image_1.jpg tests/images/image_2.jpg