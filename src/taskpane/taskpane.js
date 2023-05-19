/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* global document, Office */

const tenor_api_key = "AIzaSyAc2OphGysCfd2YVwWlIDd73yPzWJqGflM";
const giphy_api_key = "Z81NqORdTMOb6Qhsm1PD4a2pzFiHOM0X";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("q").onchange = grab_data_from_tenor_luncher;
  }
});

function insertImageLuncher(event) {
  if (event.target.src.startsWith("https://")) {
    insertImage(event.target.alt);
  } else {
    insertImage();
  }
}

function grab_data_from_tenor_luncher() {
  grab_data_from_tenor(document.getElementById("q").value);
}

function insertImage(url = "https://media.giphy.com/media/3oEjI6SIIHBdRxXI40/giphy.gif") {
  console.log("url: " + url);
  convertImageToBase64FromURL(url)
    .then((base64Image) => {
      insertImageToPowerPoint(base64Image.split(",")[1]);
    })
    .catch((err) => console.log(err));
}

function insertImageToPowerPoint(base64String) {
  Office.context.document.setSelectedDataAsync(
    base64String,
    {
      coercionType: Office.CoercionType.Image,
      imageLeft: 50,
      imageTop: 50,
      imageWidth: 400,
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
      }
    }
  );
}

function convertImageToBase64FromURL(url) {
  return fetch(url)
    .then((response) => response.blob())
    .then((blob) => {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => {
          const base64String = reader.result;
          resolve(base64String);
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
    });
}

// url Async requesting function
function httpGetAsync(theUrl, callback) {
  // create the request object
  var xmlHttp = new XMLHttpRequest();

  // set the state change callback to capture when the response comes in
  xmlHttp.onreadystatechange = function () {
    if (xmlHttp.readyState == 4 && xmlHttp.status == 200) {
      callback(xmlHttp.responseText);
    }
  };

  // open as a GET call, pass in the url and set async = True
  xmlHttp.open("GET", theUrl, true);

  // call send with no params as they were passed in on the url string
  xmlHttp.send(null);

  return;
}

// callback for the top 8 GIFs of search
function tenorCallback_search(responsetext) {
  // Parse the JSON response
  var response_objects = JSON.parse(responsetext);

  var top_10_gifs = response_objects["results"];

  // load the GIFs -- for our example we will load the first GIFs preview size (nanogif) and share size (gif)

  document.getElementById("preview_gif_0").src = top_10_gifs[0]["media_formats"]["nanogif"]["url"];
  document.getElementById("preview_gif_0").alt = top_10_gifs[0]["media_formats"]["gif"]["url"];
  document.getElementById("preview_gif_0").onclick = insertImageLuncher;

  for (var i = 1; i < top_10_gifs.length - 1; i++) {
    var img_tag = document.createElement("img");
    img_tag.id = "preview_gif_" + i;
    var container = document.getElementById("tenor-gifs");
    container.appendChild(img_tag);
    document.getElementById("preview_gif_" + i).src = top_10_gifs[i]["media_formats"]["nanogif"]["url"];
    document.getElementById("preview_gif_" + i).alt = top_10_gifs[i]["media_formats"]["gif"]["url"];
    document.getElementById("preview_gif_" + i).onclick = insertImageLuncher;
    document.getElementById("preview_gif_" + i).className = "gif-preview";
  }

  return;
}

// function to call the trending and category endpoints
function grab_data_from_tenor(search_term = "money") {
  console.log("search_term: " + search_term);
  // set the apikey and limit
  var apikey = tenor_api_key;
  var clientkey = "PowerGIF";
  var lmt = 10;

  // using default locale of en_US
  var search_url =
    "https://tenor.googleapis.com/v2/search?q=" +
    search_term +
    "&key=" +
    apikey +
    "&client_key=" +
    clientkey +
    "&limit=" +
    lmt;

  httpGetAsync(search_url, tenorCallback_search);

  // data will be loaded by each call's callback
  return;
}

grab_data_from_tenor();
