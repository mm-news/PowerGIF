/* eslint-disable no-undef */
/* global document, Office */

const tenor_api_key = "AIzaSyAc2OphGysCfd2YVwWlIDd73yPzWJqGflM";
const giphy_api_key = "Z81NqORdTMOb6Qhsm1PD4a2pzFiHOM0X";

var limit = 5;

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
    limit += 5;
    grab_gif(document.getElementById("q").value, limit);
  }
});

function isAtBottom() {
  var windowHeight = window.innerHeight;

  var documentHeight = document.documentElement.scrollHeight;

  var scrollTop = window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop;

  return scrollTop + windowHeight >= documentHeight;
}

function grab_gif_luncher() {
  limit = 10;
  grab_gif(document.getElementById("q").value);
  document.getElementById("loading-gifs").className = "";
  document.getElementById("no-more-gifs").className = "no-display";
}

async function grab_gif(q = "money", lmt = limit) {
  var tenor_response_objects = [];
  var giphy_response_objects = [];

  try {
    tenor_response_objects = await grab_data_from_tenor(q, lmt);
  } catch (error) {
    console.log(error);
  }

  giphy_response_objects = await grab_data_from_giphy(q, lmt);

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
  document.getElementById("gifs").innerHTML = "";

  if (tenor.length < limit && giphy.length < limit) {
    document.getElementById("loading-gifs").className = "no-display";
    document.getElementById("no-more-gifs").className = "";
  }
  for (var t = (g = 0); t < tenor.length || t < giphy.length; t++, g++) {
    if (t < tenor.length) {
      var img_tag_t = document.createElement("img");
      var container_t = document.getElementById("gifs");
      container_t.appendChild(img_tag_t);
      img_tag_t.src = tenor[t]["media_formats"]["nanogif"]["url"];
      img_tag_t.alt = tenor[t]["media_formats"]["gif"]["url"];
      img_tag_t.onclick = insertImageLuncher;
      img_tag_t.className = "gif-preview";
    }
    if (g < giphy.length) {
      var img_tag_g = document.createElement("img");
      var container_g = document.getElementById("gifs");

      container_g.appendChild(img_tag_g);

      img_tag_g.src = giphy[g]["images"]["fixed_height"]["url"];
      img_tag_g.alt = giphy[g]["images"]["original"]["url"];
      img_tag_g.onclick = insertImageLuncher;
      img_tag_g.className = "gif-preview";
    }
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

function grab_data_from_tenor(search_term = "money", lmt = limit) {
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

function grab_data_from_giphy(q, lmt = limit) {
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
