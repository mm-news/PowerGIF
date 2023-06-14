/* eslint-disable no-undef */
/* global document, Office */

const tenor_api_key = "AIzaSyAc2OphGysCfd2YVwWlIDd73yPzWJqGflM";
const giphy_api_key = "Z81NqORdTMOb6Qhsm1PD4a2pzFiHOM0X";

var limit = 5;
var source_tenor = true;
var source_giphy = true;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("q").onchange = grab_gif_luncher;
    document.getElementById("search-button").onclick = grab_gif_luncher;
    document.getElementById("choose-source").className = "source-no-display";
    document.getElementById("filter-button").onclick = show_chooser;
    document.getElementById("x-mark").onclick = hide_chooser;
    document.getElementById("source-tenor").onclick = set_source_tenor;
    document.getElementById("source-giphy").onclick = set_source_giphy;
  }
});

function set_source_tenor() {
  if (document.getElementById("source-tenor").checked) {
    source_tenor = true;
    document.getElementById("source-giphy").disabled = false;
  } else {
    if (source_giphy == false) {
      document.getElementById("source-giphy").checked = true;
      document.getElementById("source-giphy").disabled = true;
      source_giphy = true;
      source_tenor = false;
    } else {
      source_tenor = false;
    }
  }
  grab_gif_luncher();
}

function set_source_giphy() {
  if (document.getElementById("source-giphy").checked) {
    source_giphy = true;
    document.getElementById("source-tenor").disabled = false;
  } else {
    if (source_tenor == false) {
      document.getElementById("source-tenor").checked = true;
      document.getElementById("source-tenor").disabled = true;
      source_tenor = true;
      source_giphy = false;
    } else {
      source_giphy = false;
    }
  }
  grab_gif_luncher();
}

function show_chooser() {
  if (document.getElementById("choose-source").className === "source-no-display") {
    document.getElementById("choose-source").className = "";
    document.getElementById("header").className = "header-with-source-chooser";
    document.getElementById("content").className = "content-with-source-chooser";
    document.getElementById("filter-button").onclick = hide_chooser;
  }
}

function hide_chooser() {
  if (document.getElementById("choose-source").className === "") {
    document.getElementById("choose-source").className = "source-no-display";
    document.getElementById("header").className = "header";
    document.getElementById("content").className = "content";
    document.getElementById("filter-button").onclick = show_chooser;
  }
}

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
  limit = 5;
  grab_gif(document.getElementById("q").value);
  document.getElementById("loading-gifs").className = "";
  document.getElementById("no-more-gifs").className = "no-display";
  window.location.href = "#header";
}

async function grab_gif(q = "money", lmt = limit) {
  var tenor_response_objects = [];
  var giphy_response_objects = [];

  if (source_tenor) {
    try {
      tenor_response_objects = await grab_data_from_tenor(q, lmt);
    } catch (error) {
      console.log(error);
    }
  }

  if (source_giphy) {
    giphy_response_objects = await grab_data_from_giphy(q, lmt);
  }

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
  } else {
    document.getElementById("loading-gifs").className = "";
    document.getElementById("no-more-gifs").className = "no-display";
  }

  console.log("tenor length: " + tenor.length); // debug
  console.log("giphy length: " + giphy.length); // debug

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
    console.log("search_term: " + search_term);
    console.log("lmt: " + lmt);
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
