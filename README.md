# PPTX Image Extractor

This is a quick node script to extract images from slides; relies on having title fields which separate the other slides.

## Installation

To get setup, clone this repo and run `npm install`

This will download all the necessary dependencies.

## Pre-req before running the scripts

You'll have wanted to convert a PPTX file into an XML output. I think I found a command on my Mac to do this....but don't quote me on this.
It outputs the PPTX into a `ppt` folder that includes slides and the images. Move these into the top level of this repo to then get the scripts to run against them.

## Running the scripts

### Bulk PPT Extractor

If you have a number of PPT presentations, each is for an individual seal; use this script.

Put your PPTs in the `ppts/to-process` folder and run `npm run bulk-ppt-extractor`.
This will create a separate folder for each PPT using the filename as the folder with all the images from that PPT into the folder. The folders will be created in the `extracted-ppt-output` folder.

Be aware, if you re-run the script it will remove anything in the output `extracted-ppt-output` folder.

### Rotate images

Once you have extracted the images from the Bulk PPT Extractor, you can run a script that will copy the images into the `rotated-images/` folder and go through them all and rotate each image - 10 degrees to the left, and 10 degrees to the right - giving you 21 images per seal.

This can be run by running `node rotate-images.js`

### Other scripts

Checkout the `package.json`. The other scripts are:

- `npm run rename-images` : this will rename images to make sure they are numerically ordered
- `npm run rename-slides` : this will rename slides to make sure they are numerically ordered
- `npm run extract-text` : this extracts the text from all slides so we know what's in them
- `npm run find-title-slides` : this loops through all the slides and logs all slides where there is text and no images found in the slide
- `npm run process-slides` : this will process all the slides and output the images from each slide into a relevant folder based on the title slide
