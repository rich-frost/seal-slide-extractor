const Jimp = require("jimp");
const execSync = require("child_process").execSync;
// In newer Node.js versions where process is already global this isn't necessary.
const rimraf = require("rimraf");
const fs = require("fs");

const copyOriginalFiles = () => {
  rimraf.sync("./rotated-images/");
  fs.mkdirSync("./rotated-images/");
  execSync("cp -r ./extracted-ppt-output/ ./rotated-images/");
};

copyOriginalFiles();

const goThroughFolder = folder => {
  fs.readdir("./rotated-images/" + folder, function(err, images) {
    images.forEach(function(image, index) {
      console.log(
        "About to rotate",
        "./rotated-images/" + folder + "/" + image
      );
      const splitImage = image.split(".");
      const path = "./rotated-images/" + folder + "/" + splitImage[0];
      const ext = splitImage[1];
      readImage(path, ext).then(image => {
        rotateImage(path, ext, image, 1);
        rotateImage(path, ext, image, 350);
      });
    });
  });
};

const findFolders = () => {
  fs.readdir("./rotated-images/", function(err, folders) {
    folders.forEach(function(folder, index) {
      if (folder.indexOf(".") === -1) {
        goThroughFolder(folder);
      }
    });
  });
};

findFolders();

const rotateImage = (path, ext, image, rotation) => {
  const image1 = image.clone();
  image1
    .rotate(rotation)
    .write(
      path + "-" + rotation + "." + ext,
      checkLoop(path, ext, image, rotation)
    );
};

const checkLoop = (path, ext, image, rotation) => {
  if (rotation > 300 && rotation < 360) {
    rotation += 1;
    rotateImage(path, ext, image, rotation);
  } else if (rotation < 10) {
    rotation += 1;
    rotateImage(path, ext, image, rotation);
  }
};
const readImage = (path, ext) => {
  return Jimp.read(path + "." + ext);
};
