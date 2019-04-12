/*
Bulk PPT Extractor

The purpose of this script is to run through a folder that contains multiple PPTs.
Each PPT is expected to be for a single Seal with the name of that seal in the
PPT filename.
The script will loop through each PPT, zip it then unzip it to extract the slide
metadata and all the images.

It will then take the images and put them in a directory with the seal's name.

*/

const fs = require("fs");
const execSync = require("child_process").execSync;
// In newer Node.js versions where process is already global this isn't necessary.
const rimraf = require("rimraf");

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
  extractImages(filename);
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

const extractImages = filename => {
  execSync(
    "mv " + outputDir + filename + "/ppt/media/* " + outputDir + filename
  );
  rimraf.sync(outputDir + filename + "/_rels");
  rimraf.sync(outputDir + filename + "/docProps");
  rimraf.sync(outputDir + filename + "/ppt");
  rimraf.sync(outputDir + filename + "/" + filename + ".zip");
  rimraf.sync(outputDir + filename + "/[Content_Types].xml");
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
