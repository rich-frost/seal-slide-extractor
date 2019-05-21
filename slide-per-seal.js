/*
Slide per Seal Script

This script expects every slide to contain multiple images for 1 seal only.
Within the slide, it should contain the text of the name of the seal.

The script will loop through each slide, extract the text and create a folder for
each seal containing the images. How it should work:

- Zip, then unzip the PPT
- Loop through each slide:
  - Read the text from the slide to create the folder in the output dir
  - Work out how many images in the slide, and take the first x images from the
    unzipping process and put in the relevant folder
*/

const fs = require("fs");
const fsExtra = require("fs-extra");
// In newer Node.js versions where process is already global this isn't necessary.
const parser = require("xml2json");
const rimraf = require("rimraf");
const execSync = require("child_process").execSync;

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
const renameSlides = dir => {
  fs.readdir(dir, function(err, files) {
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
const renameImages = dir => {
  fs.readdir(dir, function(err, files) {
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

            console.warn(file, "Can't find any photos, not sure what to do");
          }
        }
      });
    });
  });
};

const toProcessDir = "./ppts/to-process/";
const processedDir = "./ppts/processed/";

const outputDir = "./extracted-ppt-output/";

rimraf.sync(outputDir);
fs.mkdirSync(outputDir);

const processPPTs = () => {
  fs.readdir(toProcessDir, function(err, files) {
    files.forEach(function(file, index) {
      let filename = file.split(".")[0];
      processPPT(filename);
    });
  });
};

const processPPT = filename => {
  //Create a new folder for the seal
  fs.mkdirSync(outputDir + filename);

  zip(filename);
  unzip(filename);
  renameSlides(outputDir + filename);
  renameImages(outputDir + filename);
  processSlides(filename);
  movePPTAfterProcess(filename);
};

const zip = filename => {
  execSync(
    "cd " +
      toProcessDir +
      " && cp " +
      filename +
      ".pptx ../." +
      outputDir +
      filename +
      "/" +
      filename +
      ".zip"
  );
};

const unzip = filename => {
  execSync("cd " + outputDir + filename + " && unzip " + filename + ".zip");
};

const processSlides = filename => {
  const slidesDir = outputDir + filename + "/ppt/slides/";
  fs.readdir(slidesDir, function(err, files) {
    files.forEach(function(file, index) {
      fs.readFile(slidesDir + file, "utf8", function(err, data) {
        console.log(data);
    });
  });
  // execSync(
  //   "mv " + outputDir + filename + "/ppt/slides/* " + outputDir + filename
  // );

  // execSync(
  //   "mv " + outputDir + filename + "/ppt/media/* " + outputDir + filename
  // );
  // rimraf.sync(outputDir + filename + "/_rels");
  // rimraf.sync(outputDir + filename + "/docProps");
  // rimraf.sync(outputDir + filename + "/ppt");
  // rimraf.sync(outputDir + filename + "/" + filename + ".zip");
  // rimraf.sync(outputDir + filename + "/[Content_Types].xml");
};

const movePPTAfterProcess = filename => {
  execSync(
    "mv " +
      toProcessDir +
      filename +
      ".pptx " +
      processedDir +
      filename +
      ".pptx "
  );
};

module.exports = {
  processPPTs
};
