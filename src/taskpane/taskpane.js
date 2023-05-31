/* eslint-disable no-undef */
/* global document, Office */

const tenor_api_key = "AIzaSyAc2OphGysCfd2YVwWlIDd73yPzWJqGflM";
const giphy_api_key = "Z81NqORdTMOb6Qhsm1PD4a2pzFiHOM0X";

var limit_t = 10;
var limit_g = 10;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("q").onchange = grab_gif_luncher;
  }
});

function insertImageLuncher(event) {
  if (event.target.src.startsWith("https://")) {
    insertImage(event.target.alt);
  } else {
    insertImage();
  }
}

window.addEventListener("scroll", function () {
  if (isAtBottom()) {
    limit_t += 10;
    limit_g += 10;
    grab_gif(document.getElementById("q").value, limit_g, limit_t);
  }
});

function isAtBottom() {
  var windowHeight = window.innerHeight;

  var documentHeight = document.documentElement.scrollHeight;

  var scrollTop = window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop;

  return scrollTop + windowHeight >= documentHeight;
}

function grab_gif_luncher() {
  limit_t = limit_g = 10;
  grab_gif(document.getElementById("q").value);
  document.getElementById("loading-gifs").className = "";
  document.getElementById("no-more-gifs").className = "no-display";
}

async function grab_gif(q = "money", lmt_g = limit_g, lmt_t = limit_t) {
  var tenor_response_objects = [];
  var giphy_response_objects = [];

  try {
    tenor_response_objects = await grab_data_from_tenor(q, lmt_t);
  } catch (error) {
    console.log(error);
  }

  giphy_response_objects = await grab_data_from_giphy(q, lmt_g);

  insertGIFtoHtml(tenor_response_objects, giphy_response_objects);
}

function insertImage(url = "https://media.giphy.com/media/3oEjI6SIIHBdRxXI40/giphy.gif") {
  console.log("url: " + url);
  convertImageToBase64FromURL(url)
    .then((base64Image) => {
      insertImageToPowerPoint(base64Image.split(",")[1]);
    })
    .catch((err) => console.log(err));
}

function insertGIFtoHtml(tenor, giphy) {
  document.getElementById("tenor-gifs").innerHTML = "";

  if (tenor.length < limit_g) {
    document.getElementById("loading-gifs").className = "no-display";
    document.getElementById("no-more-gifs").className = "";
  }
  for (var i = 0; i < tenor.length - 1; i++) {
    var img_tag_t = document.createElement("img");
    var container_t = document.getElementById("tenor-gifs");
    container_t.appendChild(img_tag_t);
    img_tag_t.src = tenor[i]["media_formats"]["nanogif"]["url"];
    img_tag_t.alt = tenor[i]["media_formats"]["gif"]["url"];
    img_tag_t.onclick = insertImageLuncher;
    img_tag_t.className = "gif-preview";
  }

  document.getElementById("giphy-gifs").innerHTML = "";

  if (giphy.length < limit_g) {
    document.getElementById("loading-gifs").className = "no-display";
    document.getElementById("no-more-gifs").className = "";
  }
  for (var j = 0; j < giphy.length - 1; j++) {
    var img_tag_g = document.createElement("img");
    var container_g = document.getElementById("giphy-gifs");

    container_g.appendChild(img_tag_g);

    img_tag_g.src = giphy[j]["images"]["fixed_height"]["url"];
    console.log(giphy[j]["images"]["fixed_height"]["url"]); //debug
    img_tag_g.alt = giphy[j]["images"]["original"]["url"];
    console.log(giphy[j]["images"]["original"]["url"]); //debug
    img_tag_g.onclick = insertImageLuncher;
    img_tag_g.className = "gif-preview";
  }

  return;
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

function grab_data_from_tenor(search_term = "money", lmt = limit_t) {
  return new Promise(function (resolve, reject) {
    console.log("search_term: " + search_term); //debug
    console.log("lmt: " + lmt); //debug
    // set the apikey
    var apikey = tenor_api_key;
    var clientkey = "PowerGIF";

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

    // send the http request
    var http = new XMLHttpRequest();
    http.open("GET", search_url, true);
    http.onload = function () {
      if (http.readyState === 4 && http.status === 200) {
        var response = JSON.parse(http.responseText);
        var tenor_response_objects = response["results"];
        resolve(tenor_response_objects);
      } else {
        reject("Error: " + http.status);
      }
    };
    http.send(null);
  });
}

function grab_data_from_giphy(q, lmt = limit_g) {
  if (lmt >= 50) {
    lmt = 49; //Giphy API limit
  }
  const url = `https://api.giphy.com/v1/gifs/search?api_key=${giphy_api_key}&q=${q}&limit=${lmt}`;

  return fetch(url)
    .then((response) => response.json())
    .then((data) => data.data)
    .catch((error) => console.log(error));
}

grab_gif();
