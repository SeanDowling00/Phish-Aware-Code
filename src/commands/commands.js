Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});
function loadWebPage(url, elementId) {
  const xhr = new XMLHttpRequest();
  xhr.onreadystatechange = function () {
    if (this.readyState === 4 && this.status === 200) {
      document.getElementById(elementId).innerHTML = xhr.responseText;
    }
  };
  xhr.open("GET", url, true);
  xhr.send();
}

function getParameterByName(url, name) {
  name = name.replace(/[\[\]]/g, "\\$&");
  var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
    results = regex.exec(url);
  if (!results) return null;
  if (!results[2]) return "";
  return results[2].replace(/\+/g, " ");
}

export async function run() {
  const item = Office.context.mailbox.item;

  // Get the attachments collection
  const attachments = item.attachments;

  Office.context.mailbox.item.getAllInternetHeadersAsync(function (result) {
    if (result.status === "succeeded") {
      var spf = result.value.match(/spf=([^;\s]+)/g);
      if (spf) {
        spf = spf[0].substring(4);
      } else {
        spf = "Not Found";
      }
      var dkim = result.value.match(/dkim=([^;\s]+)/g);
      if (dkim) {
        dkim = dkim[0].substring(5);
      } else {
        dkim = "Not Found";
      }
      var dmarc = result.value.match(/dmarc=([^;\s]+)/g);
      if (dmarc) {
        dmarc = dmarc[0].substring(6);
      } else {
        dmarc = "Not Found";
      }

      document.getElementById("item-results").classList.remove("fraud");
      document.getElementById("item-results").classList.remove("authentic");

      if (
        spf === "Not Found" ||
        dkim === "Not Found" ||
        dmarc === "Not Found" ||
        spf === "none" ||
        dkim === "none" ||
        dmarc === "none" ||
        spf === "fail" ||
        dkim === "fail" ||
        dmarc === "fail"
      ) {
        var results = "Potentially Fraudulent";
        document.getElementById("item-results").classList.add("fraud");
        var fraudInfo =
          "This email may be fraudulant, this is either because <br/> A) The analysis returned failed results <br/> B) The checks could not be complete";
        document.getElementById("item-fraudInfo").classList.add("fraud");
        var proceed =
          "This analysis does not mean this email is 100% fraudulant or authentic, spoofing attacks can still occur!!";
        document.getElementById("item-proceed").classList.add("fraud");
      } else {
        var results = "Authentic Email";
        document.getElementById("item-results").classList.add("authentic");
        var fraudInfo =
          "This is more than likely an authentic Email, this means the email may <br/> A) Not be Spoofed <br/> B) Not be a Phishing Email";
        document.getElementById("item-fraudInfo").classList.add("authentic");
        var proceed =
          "This analysis does not mean this email is 100% fraudulant or authentic, spoofing attacks can still occur!";
        document.getElementById("item-proceed").classList.add("fraud");
      }
      var info = "Click each button below for more information";

      document.getElementById("item-results").innerHTML = results;
      document.getElementById("item-fraudInfo").innerHTML = fraudInfo;
      document.getElementById("item-proceed").innerHTML = proceed;
      document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
      document.getElementById("item-sender").innerHTML = "<b>Sender:</b> <br/>" + item.sender.emailAddress;
      document.getElementById("item-date").innerHTML = "<b>Received:</b> <br/>" + item.dateTimeCreated;
      document.getElementById("item-dkim").innerHTML = "<b>DKIM:</b> <br/>" + dkim;
      document.getElementById("item-spf").innerHTML = "<b>SPF:</b> <br/>" + spf;
      document.getElementById("item-dmarc").innerHTML = "<b>DMARC:</b> <br/> " + dmarc;
      document.getElementById("item-info").innerHTML = info;
    } else {
      console.error(`Error getting internet headers: ${result.error.message}`);
    }
  });

  Office.context.mailbox.item.body.getAsync("text", function (result) {
    var urls = Office.context.mailbox.item.getEntities().urls;

    var filteredUrls = [];

    for (var i = 0; i < urls.length; i++) {
      var url = urls[i];
      if (url.includes("eur03.safelinks.protection.outlook.com")) {
        var decodedUrl = decodeURIComponent(getParameterByName(url, "url"));
        url = decodedUrl;
      }
      if (!url.includes("setu") && !url.includes("mailto")) {
        filteredUrls.push(url.replace(/^(?:https?:\/\/)?(?:www\.)?([^\/]+)\/?.*$/, "$1"));
      }
    }

    document.getElementById("item-links").innerHTML = "URLs found in this email: " + filteredUrls;
    // Create a parent div element to contain all the responses
    var responseContainer = document.createElement("div");
    document.body.appendChild(responseContainer);

    var scannedUrls = [];

    for (var i = 0; i < filteredUrls.length; i++) {
      if (scannedUrls.indexOf(filteredUrls[i]) === -1) {
        scannedUrls.push(filteredUrls[i]); // add scanned URL to the array
        (function (url) {
          // create closure around url variable
          var xhr = new XMLHttpRequest();
          var apiUrl = `https://api-phish-proxy.azurewebsites.net/scan?url=${encodeURIComponent(url)}`;
          xhr.open("GET", apiUrl, true);
          xhr.setRequestHeader("Content-Type", "application/json");
          xhr.setRequestHeader("Access-Control-Allow-Origin", "*");
          xhr.setRequestHeader("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
          xhr.onreadystatechange = function () {
            if (this.readyState === XMLHttpRequest.DONE && this.status === 200) {
              var responseDiv = document.createElement("div");
              responseDiv.innerHTML = `<u><b>${url}</b>: <b>${this.responseText}</b>`;
              document.getElementById("item-vtlinksresponse").appendChild(responseDiv);
            } else if (this.readyState === XMLHttpRequest.DONE) {
              console.error(`Request failed with status: ${this.status}`);
              var responseDiv = document.createElement("div");
              responseDiv.innerHTML = `Error for ${url}: Request failed with status: ${this.status}`;
              document.getElementById("item-vtlinksresponse").appendChild(responseDiv);
            }
          };
          xhr.send();
        })(filteredUrls[i]); // pass url variable into closure function
      }
    }
  });

  let attachmentInfo = "";
  if (attachments.length === 0) {
    attachmentInfo = "No attachments found.";
  } else {
    attachmentInfo = "<ul>";
    for (let i = 0; i < attachments.length; i++) {
      const attachment = attachments[i];
      let contentType = attachment.contentType;
      if (contentType.startsWith("image/")) {
        if (contentType === "image/jpeg") {
          contentType = "JPEG image";
        } else if (contentType === "image/png") {
          contentType = "PNG image";
        } else if (contentType === "image/gif") {
          contentType = "GIF image";
        } else {
          contentType = "Image";
        }
      } else if (contentType.includes("pdf")) {
        contentType = "PDF document";
      } else if (contentType.startsWith("audio/")) {
        contentType = "Audio file";
      } else if (contentType === "video/mp4") {
        contentType = "MP4 video";
      } else if (contentType.startsWith("application/vnd.openxmlformats-officedocument.")) {
        if (contentType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
          contentType = "Microsoft Word document";
        } else if (contentType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
          contentType = "Microsoft Excel spreadsheet";
        } else {
          contentType = "Office document";
        }
      } else {
        contentType = attachment.contentType;
      }
      attachmentInfo += `<li><strong>Name:</strong> ${attachment.name} <br>
                         <strong>Size:</strong> ${attachment.size} Bytes<br>
                         <strong>Content Type:</strong> ${contentType}</li>`;
      attachmentInfo += "<br>";
    }
    attachmentInfo += "</ul>";
  }

  document.getElementById("item-attachment").innerHTML = attachmentInfo;
  //    var link = "<a href='https://attachment-api.azurewebsites.net/'>Click Here</a>";
  // document.getElementById("item-attachmentresults").innerHTML = "To scan any file, " + link + " and then drag and drop the file to be scanned onto the upload button";

  // Call the function to load the webpage and insert it into an HTML element with an ID of "myDiv"
  loadWebPage("https://attachment-api.azurewebsites.net/", "item-attachmentresults");
}
