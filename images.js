/*
Images Script

This script was originally created based on a PPT which had a title slide with
the name of the seal. All following slides then contained images of that seal.

This scipt would use the different functions to extract all the images then
work out which slides were for which seal based on the location of the title slides.

*/

const fs = require("fs");
const parser = require("xml2json");
const rimraf = require("rimraf");

const baseSlideDir = "./slides/slides/";
const baseImageDir = "./slides/media/";

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
      const ext = "xml";
      fs.rename(
        baseSlideDir + file,
        baseSlideDir + "slide" + formattedNum + "." + ext,
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
  fs.readdir(baseSlideDir, function(err, files) {
    files.forEach(function(file, index) {
      fs.readFile(baseSlideDir + file, "utf8", function(err, data) {
        if (!data) return;
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

// Loop through all the slide files and output the title slides with slide number
const findTitleSlides = () => {
  let previousTitle = {};
  const files = fs.readdirSync(baseSlideDir);
  files.forEach(function(file, index) {
    const filename = baseSlideDir + file;
    if (!fs.lstatSync(filename).isDirectory()) {
      const data = fs.readFileSync(filename, "utf8");

      if (data && data.indexOf("p:pic") === -1) {
        try {
          const regex = new RegExp(/<a:t>([\s\S]*?)<\/a:t>/);
          const title = regex.exec(data)[1];

          if (Object.keys(previousTitle).length > 0) {
            console.warn(
              "Previous",
              previousTitle.title,
              previousTitle.index + 1,
              index - 1
            );
          }
          previousTitle = { title, index, file };
        } catch (e) {}
      }
    }
  });
};

// Loop through all the slide files, find the titles and work out the slides
// in between. Then process each slide xml file to extract the images referenced
// in them. Fetch them and save them to the output folders
const extractImagesFromSlides = () => {
  let previousTitle = {};
  let seals = [];
  const files = fs.readdirSync(baseSlideDir);
  files.forEach(function(file, index) {
    const filename = baseSlideDir + file;
    if (!fs.lstatSync(filename).isDirectory()) {
      const data = fs.readFileSync(filename, "utf8");

      if (data && data.indexOf("p:pic") === -1) {
        try {
          const regex = new RegExp(/<a:t>([\s\S]*?)<\/a:t>/);
          const title = regex.exec(data)[1];

          if (Object.keys(previousTitle).length > 0) {
            seals.push({
              seal: previousTitle.title,
              start: previousTitle.index + 1,
              end: index
            });
          }
          previousTitle = { title, index, file };
          //console.warn(file, title);
        } catch (e) {}
      }
    }
  });

  seals.forEach(seal => {
    getImagesForTheSeal(seal);
  });
};

const getImagesForTheSeal = ({ seal, start, end }) => {
  let index = 1;

  //Create a folder for the seal
  const sealFolder = outputDir + seal;
  if (!fs.existsSync(sealFolder)) {
    fs.mkdirSync(sealFolder);
  } else {
    fs.readdirSync(sealFolder, (err, files) => {
      index = files.length;
    });
  }

  // Read each slide res file and read the images found
  for (let i = start; i < end; i++) {
    const data = fs.readFileSync(
      "./slides/slides/slide.xml/slide" + i + ".xml.rels",
      "utf8"
    );

    const regex = /Target="..\/media([\s\S]*?)"/g;
    while ((m = regex.exec(data)) !== null) {
      // This is necessary to avoid infinite loops with zero-width matches
      if (m.index === regex.lastIndex) {
        regex.lastIndex++;
      }
      const image = m[0].replace('Target="../media/', "").replace('"', "");
      const file = "./slides/media/" + image;

      if (fs.existsSync(file)) {
        const extDot = file.lastIndexOf(".");
        const ext = file.substr(extDot);
        fs.renameSync(file, sealFolder + "/" + seal + "-" + index + ext);
        console.log(
          i,
          index,
          seal,
          file,
          sealFolder + "/" + seal + "-" + index + ext
        );

        index += 1;
      }
    }
  }
};

module.exports = {
  renameSlides,
  renameImages,
  extractSlideText,
  extractImagesFromSlides,
  findTitleSlides
};
