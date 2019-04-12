/*
Images Script

This script was originally created based on a PPT which had a title slide with
the name of the seal. All following slides then contained images of that seal.

This scipt would use the different functions to extract all the images then
work out which slides were for which seal based on the location of the title slides.

It could do with going over again, or re-written as I can't remember the exact
functionality used on the PPT that was 8GB.
*/

const fs = require("fs");
const fsExtra = require("fs-extra");
const path = require("path");
// In newer Node.js versions where process is already global this isn't necessary.
const process = require("process");
const parser = require("xml2json");
const rimraf = require("rimraf");

const baseSlideDir = "./slides/slides/";
const baseImageDir = "./slides/images/";

const slideDir = "./slides/formatted-slides/";
const imagesDir = "./slides/formatted-images/";
const outputDir = "./output/";

rimraf.sync(outputDir);
fs.mkdirSync(outputDir);

const convertNumber = file => {
  const number = file.substr(5).split(".")[0];
  const ext = file.split(".")[1];
  let formattedNum = number;

  if (number.length === 1) {
    formattedNum = "000" + number;
  } else if (number.length === 2) {
    formattedNum = "00" + number;
  } else if (number.length === 3) {
    formattedNum = "0" + number;
  }

  return formattedNum;
};

// Loop through all the slide files
const renameSlides = () => {
  fs.readdir(baseSlideDir, function(err, files) {
    files.forEach(function(file, index) {
      const formattedNum = convertNumber(file);

      fs.rename(
        baseSlideDir + file,
        slideDir + "slide" + formattedNum + "." + ext,
        function(err) {
          if (err) console.log("ERROR: " + err);
        }
      );
    });
  });
};

// Loop through all the image files
const renameImages = () => {
  fs.readdir(baseImageDir, function(err, files) {
    files.forEach(function(file, index) {
      const formattedNum = convertNumber(file);

      fs.createReadStream(baseImageDir + file).pipe(
        fs.createWriteStream(imagesDir + "image" + formattedNum + "." + ext)
      );
    });
  });
};

// Loop through all the slide files - this will output all text; some are maybe wrong
const extractSlideText = () => {
  fs.readdir(slideDir, function(err, files) {
    files.forEach(function(file, index) {
      fs.readFile(slideDir + file, "utf8", function(err, data) {
        const json = JSON.parse(parser.toJson(data));

        try {
          const title =
            json["p:sld"]["p:cSld"]["p:spTree"]["p:sp"][0]["p:txBody"]["a:p"][
              "a:r"
            ]["a:t"];

          if (title.indexOf("is") > -1) {
            console.warn(file, title);
          }
        } catch (e) {}
      });
    });
  });
};

// Loop through all the slide files - this will output all text; some are maybe wrong
const collateImagesFromSlides = () => {
  fs.readdir(slideDir, function(err, files) {
    let currentSeal = "";
    let picCounter = 0;

    files.forEach(function(file, index) {
      fs.readFile(slideDir + file, "utf8", function(err, data) {
        const json = JSON.parse(parser.toJson(data));

        try {
          // Work out if this is a title slide

          const title =
            json["p:sld"]["p:cSld"]["p:spTree"]["p:sp"][0]["p:txBody"]["a:p"][
              "a:r"
            ]["a:t"];
          const pics = json["p:sld"]["p:cSld"]["p:spTree"]["p:pic"];

          // If there are no pics then we'll assume this is a definitive title slide
          if (!pics) {
            console.warn(file, title);
            currentSeal = title.replace(" ", "_");
          }
          try {
            fs.mkdirSync(outputDir + currentSeal);
          } catch (e) {}
        } catch (e) {
          if (file === "slide0653.xml") {
            console.warn(e);
          }
          // Is not a title slide, so must show seals
          try {
            const pics = json["p:sld"]["p:cSld"]["p:spTree"]["p:pic"];

            const imageFiles = fs.readdirSync(imagesDir);

            for (let i = 0; i < pics.length; i++) {
              fsExtra.copySync(
                imagesDir + imageFiles[picCounter],
                "./output/" + currentSeal + "/" + imageFiles[picCounter]
              );
              picCounter++;
            }
          } catch (e) {
            if (file === "slide0653.xml") {
              console.warn(e);
            }

            console.warn(file, "cant find any photos, not sure what to do");
          }
        }
      });
    });
  });
};

module.exports = {
  renameSlides,
  renameImages,
  extractSlideText,
  collateImagesFromSlides
};
