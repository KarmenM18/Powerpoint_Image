// reference: https://github.com/bensonruan/webcam-easy

//import Webcam from 'webcam-easy';

// Declare constants using webcam-easy library to interact with webcam
const webcamElement = document.getElementById('webcam');
const canvasElement = document.getElementById('canvas');
const snapSoundElement = document.getElementById('snapSound');
const webcam = new Webcam(webcamElement, 'user', canvasElement, snapSoundElement);

var globalImage = new Image();

  // Callback check to ensure that add-in code runs in sync with Powerpoint application's state
  Office.onReady(function (info) {
    if (info.host === Office.HostType.PowerPoint) {
      document.getElementById("takePicture").onclick = function () { // onClick event handler assigned to button, executes picture capture
        takePicture();
      };
    }
  });

async function takePictureAndInsert() {
  try {

    // Capture image from webcam
    const capturedImage = await captureImage();

    // Insert the image path into the presentation
    insertImageIntoPresentation(imagePath);

  } catch (error) {
    console.error('Error capturing image:', error);
  }
}

// PURPOSE: Take picture through webcam, stores picture data
function captureImage() {

  // Initialize the webcam instance (assuming you have the necessary HTML elements)
  const webcam = new Webcam(document.getElementById('webcam'), 'user', document.getElementById('canvas'), document.getElementById('snapSound'));

  // Engage webcam, take picture 
  webcam.start()
    .then(result => {
      console.log("webcam started");
    })
    .catch(err => {
      console.log(err);
    });

  // Take photo
  var picture = webcam.snap();
  webcam.stop();
  return picture; // note: in webcam-easy, snap() returns 'data,' which should be a PNG URL

}

// PURPOSE: takes the Image URL obtained from webcam capture, inserts into powerpoint
function insertImageIntoPresentation(imageURL) {

  // Use Office API to set ImageURL data asynchronously
  Office.context.document.setSelectedDataAsync(
    imageURL,
    {
      // Define properties for the image to be inserted
      coercionType: Office.CoercionType.Image,
      imageLeft: 0,  // Set the left position of the image
      imageTop: 0,   // Set the top position of the image
      imageWidth: 400,  // Set the width of the image
      imageHeight: 300  // Set the height of the image
    },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log('Image inserted into PowerPoint');
      } else {
        console.error('Error inserting image into PowerPoint:', result.error.message);
      }
    }
  );
}













