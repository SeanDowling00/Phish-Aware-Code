Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

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

      var internetHeaders = result.value.split(/\r\n/);
      for (var i = 0; i < internetHeaders.length; i++) {
        console.log(internetHeaders[i]);
      }

      var toList = item.to;
      var emailList = toList.map(function (to) {
        return to.emailAddress;
      });

      var returnPath = result.value.match(/Return-Path:\s*([^;\s]+)/i);
      if (returnPath) {
        returnPath = returnPath[1];
      } else {
        returnPath = "Not Found";
      }

      document.getElementById("item-results").classList.remove("fraud");
      document.getElementById("item-results").classList.remove("authentic");

      if (
        spf === "Not Found" ||
        dkim === "Not Found" ||
        dmarc === "Not Found" ||
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
      document.getElementById("item-to").innerHTML = "<b>To:</b> <br/>" + emailList.join(", ");
      var senderMatch = item.sender.emailAddress === returnPath;
      var senderElement = document.getElementById("item-sender");
      var returnElement = document.getElementById("item-return");

      if (senderMatch) {
        senderElement.innerHTML =
          "<b>Sender:</b><br/>(Should match with 'Return Path')<br/>" +
          item.sender.emailAddress +
          " <span style='color:green'>&#10004;</span>";
        returnElement.innerHTML =
          "<b>Return Path:</b><br/>" + returnPath + " <span style='color:green'>&#10004;</span>";
      } else {
        senderElement.innerHTML =
          "<b>Sender:</b><br/>(Should match with 'Return Path')<br/>" +
          item.sender.emailAddress +
          " <span style='color:red'>&#10006;</span>";
        returnElement.innerHTML = "<b>Return Path:</b><br/>" + returnPath + " <span style='color:red'>&#10006;</span>";
      }

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
    var scannedUrls = [];
    var newUrls = [];
    var excludedUrls = [];
    for (var i = 0; i < urls.length; i++) {
      var url = urls[i];
      if (url.includes("eur03.safelinks.protection.outlook.com")) {
        var decodedUrl = decodeURIComponent(getParameterByName(url, "url"));
        url = decodedUrl;
      }
      if (!url.includes("setu") && !url.includes("mailto")) {
        if (
          !url.includes("facebook.com") &&
          !url.includes("instagram.com") &&
          !url.includes("twitter.com") &&
          !url.includes("youtube.com") &&
          !url.includes("tiktok.com") &&
          !url.includes("google.com") &&
          !url.includes("itcsu.ie")
        ) {
          if (!url.includes("bit.ly")) {
            var filteredUrl = url.replace(/^(?:https?:\/\/)?(?:www\.)?([^\/]+)\/?.*$/, "$1");
            if (filteredUrls.indexOf(filteredUrl) < 0) {
              // Check if URL already exists in array
              filteredUrls.push(filteredUrl);
              var newUrl = `https://api-phish-proxy.azurewebsites.net/scan?url=${encodeURIComponent(filteredUrl)}`;
              newUrls.push(newUrl);
            }
          } else {
            // Append message to user if bit.ly link is found
            document.getElementById("item-links").innerHTML +=
              "<b>A 'bit.ly' URL was found, these need to manually scanned on <a href=https://www.virustotal.com/gui/home/search>Virustotal</a> or <a href=https://urlscan.io/>URL Scan</a><br><br>";
          }
        } else {
          var excludedUrl = url.replace(/^(?:https?:\/\/)?(?:www\.)?([^\/]+)\/?.*$/, "$1");
          excludedUrls.push(excludedUrl); // Add excluded URLs to the new array
        }
      }
    }

    if (urls.length == 0) {
      document.getElementById("item-links").innerHTML = "No URLs found";
      document.getElementById("item-vtlinksresponse").innerHTML = "Please continue your analysis";
    }

    if (filteredUrls.length > 0) {
      document.getElementById("item-links").innerHTML += "Domains found in this email: <br><ul>";
      for (var i = 0; i < filteredUrls.length; i++) {
        document.getElementById("item-links").innerHTML += "<li>" + filteredUrls[i] + "</li>";
      }
      document.getElementById("item-links").innerHTML += "</ul>";
    }
    if (excludedUrls.length > 0) {
      document.getElementById("item-links").innerHTML +=
        "<br>Large corporation domains will not be scanned by API Scanners, that being said, below is a list of domains that were not scanned:<br><ul>";
      for (var i = 0; i < excludedUrls.length; i++) {
        document.getElementById("item-links").innerHTML += "<li>" + excludedUrls[i] + "</li>";
      }
      document.getElementById("item-links").innerHTML += "</ul>";
    }

    var responseContainer = document.createElement("div");
    document.body.appendChild(responseContainer);

    for (var i = 0; i < filteredUrls.length; i++) {
      if (scannedUrls.indexOf(filteredUrls[i]) === -1) {
        scannedUrls.push(filteredUrls[i]);
        (function (url) {
          var xhr = new XMLHttpRequest();
          var apiUrl = `https://api-phish-proxy.azurewebsites.net/scan?url=${encodeURIComponent(url)}`;
          xhr.open("GET", apiUrl, true);
          xhr.setRequestHeader("Content-Type", "application/json");
          xhr.setRequestHeader("Access-Control-Allow-Origin", "*");
          xhr.setRequestHeader("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
          xhr.onreadystatechange = function () {
            if (this.readyState === XMLHttpRequest.DONE && this.status === 200) {
              var responseDiv = document.createElement("div");
              responseDiv.innerHTML = `<div style="text-align: center;">
          <p><h3><b><u>URL:<br></b></u> ${url}</h3></p>
          </div> <b>${this.responseText}</b>`;
              document.getElementById("item-vtlinksresponse").appendChild(responseDiv);
            } else if (this.readyState === XMLHttpRequest.DONE) {
              console.error(`Request failed with status: ${this.status}`);
              var responseDiv = document.createElement("div");
              responseDiv.innerHTML = `Error for ${url}: This URL could not be scanned for some reason, please retry or go to <a href=https://api-phish-proxy.azurewebsites.net/scan?url= target="_blank">Here</a> and copy the link you want scanned and paste it to the end of the URL, if this method still does not work, please go <a href=https://www.virustotal.com/gui/home/search>Virustotal</a> or <a href=https://urlscan.io/>URL Scan</a> and copy the link to be scanned`;
              document.getElementById("item-vtlinksresponse").appendChild(responseDiv);
            }
          };
          xhr.send();
        })(filteredUrls[i]);
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
      attachmentInfo += `<li><strong>Name:</strong> ${attachment.name} </li>
                         <li><strong>Size:</strong> ${attachment.size} Bytes</li>
                         <li><strong>Content Type:</strong> ${contentType}</li>`;
      attachmentInfo += "";
    }
    attachmentInfo += "</ul>";
  }

  document.getElementById("item-attachment").innerHTML = attachmentInfo;

  async function forwardEmail() {
    const item = Office.context.mailbox.item;
    var urls = Office.context.mailbox.item.getEntities().urls;

    var filteredUrls = [];

    var scannedUrls = [];
    var newUrls = [];

    for (var i = 0; i < urls.length; i++) {
      var url = urls[i];
      if (url.includes("eur03.safelinks.protection.outlook.com")) {
        var decodedUrl = decodeURIComponent(getParameterByName(url, "url"));
        url = decodedUrl;
      }
      if (!url.includes("setu") && !url.includes("mailto")) {
        var filteredUrl = url.replace(/^(?:https?:\/\/)?(?:www\.)?([^\/]+)\/?.*$/, "$1");
        filteredUrls.push(filteredUrl);
        var newUrl = `https://api-phish-proxy.azurewebsites.net/scan?url=${encodeURIComponent(filteredUrl)}`;
        newUrls.push(newUrl);
      }
    }
    // Get email properties
    const subject = item.subject;
    const sender = item.sender.emailAddress;
    const received = item.dateTimeCreated.toLocaleString();
    const body = await new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Html, {}, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    });
    const to = JSON.stringify(item.to);

    for (var i = 0; i < newUrls.length; i++) {
      var link = document.createElement("a");
      link.href = newUrls[i];
      link.target = "_blank";
      link.innerText = newUrls[i];
      // document.body.appendChild(link);
    }

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

        var internetHeaders = result.value.split(/\r\n/);
        for (var i = 0; i < internetHeaders.length; i++) {
          console.log(internetHeaders[i]);
        }

        var toList = item.to;
        var emailList = toList.map(function (to) {
          return to.emailAddress;
        });

        var returnPath = result.value.match(/Return-Path:\s*([^;\s]+)/i);
        if (returnPath) {
          returnPath = returnPath[1];
        } else {
          returnPath = "Not Found";
        }

        // Format the email body
        const emailBody = `
      <strong>The email below was determined to be a phishing email by the user, please review and take appropriate actions</strong><br>
      <br>
      <hr>
      <br>
      <strong>To:</strong> ${emailList.join(", ")}<br>
      <strong>Sender:</strong> ${sender}<br>
      <strong>Return Path:</strong> ${returnPath}<br>
      <strong>Received:</strong> ${received}<br>
      <strong>Subject:</strong> ${subject}<br>
      <strong>SPF:</strong> ${spf}<br>
      <strong>DMARC:</strong> ${dmarc}<br>
      <strong>DKIM:</strong> ${dkim}<br>
      <br>
      <hr>
      <br>
      <strong>Scanned URLs:</strong>
      <br>
      ${newUrls.map((url) => `<a href="${url}" target="_blank">${url}</a><br>`).join("")}
      <br>
      <hr>
      <br>
      <strong><a href="https://attachment-api.azurewebsites.net">Scan Attachments Here</a></strong>
      <br>
      <br>
      <hr>
      <br>
      <strong><u>Email Body</u></strong><br>
      <br>
      ${body}
      <br>
      <hr>
      <br>
      <strong><u>Internet Headers</u></strong><br>
      <br>
      <pre>${result.value}</pre>
      <br>
      <hr>
    `;

        // Create the new message and set the recipients and subject
        const message = Office.context.mailbox.displayNewMessageForm({
          toRecipients: ["sean.dowling185@gmail.com"],
          subject: `PHISHING ATTEMPT: (${subject})`,
          htmlBody: `${emailBody}`,
        });
      }
      Office.context.ui.displayDialogAsync(
        "https://phishdetectionandguidance.azurewebsites.net/assets/notification.html",
        { height: 30, width: 20 },
        (result) => {
          if (result.status === "succeeded") {
            const dialog = result.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (event) => {
              dialog.close();
            });
          }
        }
      );
    });
  }
  document.getElementById("forward-email").onclick = forwardEmail;
}
