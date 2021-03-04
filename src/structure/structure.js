// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

import nlp from "compromise";

let target = 100000;

/* global document, Office, Word */

const worker = new Worker("structure_worker.js");

worker.onmessage = function(e) {
  document.getElementById("debug").innerHTML = "Message received.";
  var words = e.data.words;
  document.getElementById("word-count").innerHTML = words.toLocaleString();
  displayStructure(words);
};

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = refresh;
    document.getElementById("structure").onchange = refresh;
    document.getElementById("target").onchange = onTarget;

    refresh();

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function onTarget() {
  target = document.getElementById("target").value;
  refresh();
}

function loading() {
  document.getElementById("word-count").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("structure-table").innerHTML = `<div class="ms-Spinner"></div>`;

  var SpinnerElements = document.querySelectorAll(".ms-Spinner");
  for (var i = 0; i < SpinnerElements.length; i++) {
    new fabric['Spinner'](SpinnerElements[i]);
  }
}

function refresh() {
  worker.postMessage("Hello world!");
  loading();

  Word.run(function (context) {
      let body = context.document.body;
      body.load("text");

      return context.sync()
        .then(function() {
          worker.postMessage({"text": body.text});
          document.getElementById("debug").innerHTML = "Message sent.";
        })
        .then(context.sync);
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function three_act(words, target) {
  return `<table class="ms-Table" style="width: 100%" id="structure-table">
    <thead>
      <tr>
        <th>Act</th>
        <th>Position</th>
        <th>Current Words</th>
        <th>Target Words</th>
      </tr>
    </thead>
    <tr>
      <td>One</td>
      <td>0% to 25%</td>
      <td>` + Math.round(words * 0.25).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.25).toLocaleString() + `</td>
    </tr>
    <tr>
      <td>Two</td>
      <td>25% to 75%</td>
      <td>` + Math.round(words * 0.5).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.5).toLocaleString() + `</td>
    </tr>
    <tr>
      <td>Three</td>
      <td>75% to 100%</td>
      <td>` + Math.round(words * 0.25).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.25).toLocaleString() + `</td>
    </tr>
  </table>`;
}

function four_act(words, target) {
  return `<table class="ms-Table" style="width: 100%" id="structure-table">
    <thead>
      <tr>
        <th>Act</th>
        <th>Position</th>
        <th>Current Words</th>
        <th>Target Words</th>
      </tr>
    </thead>
    <tr>
      <td>One</td>
      <td>0% to 25%</td>
      <td>` + Math.round(words * 0.25).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.25).toLocaleString() + `</td>
    </tr>
    <tr>
      <td>Two</td>
      <td>25% to 50%</td>
      <td>` + Math.round(words * 0.25).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.25).toLocaleString() + `</td>
    </tr>
    <tr>
      <td>Three</td>
      <td>50% to 75%</td>
      <td>` + Math.round(words * 0.25).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.25).toLocaleString() + `</td>
    </tr>
    <tr>
      <td>Four</td>
      <td>75% to 100%</td>
      <td>` + Math.round(words * 0.25).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.25).toLocaleString() + `</td>
    </tr>
  </table>`;
}

function save_the_cat(words, target) {
  return `Act One
  <table class="ms-Table" style="width: 100%" id="structure-table">
    <thead>
      <tr>
        <th>Beat</th>
        <th>Position</th>
        <th>Current Words</th>
        <th>Target Words</th>
      </tr>
    </thead>
    <tr>
      <td>Opening Image</td>
      <td>0% to 1%</td>
      <td>` + Math.round(words * 0.01).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.01).toLocaleString() + `</td>
    </tr>
    <tr>
      <td>Theme Stated</td>
      <td>5%</td>
      <td> - </td>
      <td> - </td>
    </tr>
    <tr>
      <td>Setup</td>
      <td>1% to 10%</td>
      <td>` + Math.round(words * 0.1).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.1).toLocaleString() + `</td>
    </tr>
    <tr>
      <td>Catalyst</td>
      <td>10%</td>
      <td> - </td>
      <td> - </td>
    </tr>
    <tr>
      <td>Debate</td>
      <td>10% to 20%</td>
      <td>` + Math.round(words * 0.1).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.1).toLocaleString() + `</td>
    </tr>
    <tr>
      <td>Break Into Two</td>
      <td>20%</td>
      <td> - </td>
      <td> - </td>
    </tr>
    <thead>
      <tr style="background-color: #d2d0ce">
        <th></th>
        <th>Total: 20%</th>
        <th>Total: ` + Math.round(words * 0.2).toLocaleString() + `</th>
        <th>Total: ` + Math.round(target * 0.2).toLocaleString() + `</th>
      </tr>
    </thead>
  </table><br><br>

  Act Two
  <table class="ms-Table" style="width: 100%" id="structure-table">
    <thead>
      <tr>
        <th>Beat</th>
        <th>Position</th>
        <th>Current Words</th>
        <th>Target Words</th>
      </tr>
    </thead>
    <tr>
      <td>B Story</td>
      <td>22%</td>
      <td> - </td>
      <td> - </td>
    </tr>
    <tr>
      <td>Fun and Games</td>
      <td>20% to 50%</td>
      <td>` + Math.round(words * 0.3).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.3).toLocaleString() + `</td>
    </tr>
    <tr>
      <td>Midpoint</td>
      <td>50%</td>
      <td> - </td>
      <td> - </td>
    </tr>
    <tr>
      <td>Bad Guys Close In</td>
      <td>50% to 75%</td>
      <td>` + Math.round(words * 0.25).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.25).toLocaleString() + `</td>
    </tr>
    <tr>
      <td>All is Lost</td>
      <td>75%</td>
      <td> - </td>
      <td> - </td>
    </tr>
    <tr>
      <td>Dark Night of the Soul</td>
      <td>75% to 80%</td>
      <td>` + Math.round(words * 0.05).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.05).toLocaleString() + `</td>
    </tr>
    <tr>
      <td>Break into Three</td>
      <td>80%</td>
      <td> - </td>
      <td> - </td>
    </tr>
    <thead>
      <tr style="background-color: #d2d0ce">
        <th></th>
        <th>Total: 60%</th>
        <th>Total: ` + Math.round(words * 0.6).toLocaleString() + `</th>
        <th>Total: ` + Math.round(target * 0.6).toLocaleString() + `</th>
      </tr>
    </thead>
  </table><br><br>

  Act Three
  <table class="ms-Table" style="width: 100%" id="structure-table">
    <thead>
      <tr>
        <th>Beat</th>
        <th>Position</th>
        <th>Current Words</th>
        <th>Target Words</th>
      </tr>
    </thead>
    <tr>
      <td>Finale</td>
      <td>80% to 99%</td>
      <td>` + Math.round(words * 0.19).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.19).toLocaleString() + `</td>
    </tr>
    <tr>
      <td>Final Image</td>
      <td>99% to 100%</td>
      <td>` + Math.round(words * 0.01).toLocaleString() + `</td>
      <td>` + Math.round(target * 0.19).toLocaleString() + `</td>
    </tr>
    <thead>
      <tr style="background-color: #d2d0ce">
        <th></th>
        <th>Total: 20%</th>
        <th>Total: ` + Math.round(words * 0.2).toLocaleString() + `</th>
        <th>Total: ` + Math.round(target * 0.2).toLocaleString() + `</th>
      </tr>
    </thead>
  </table>
  `;
}

function displayStructure(words) {
  let structure = document.getElementById("structure").value;

  switch(structure) {
    case "three-act":
      document.getElementById("structure-table").innerHTML = three_act(words, target);
      break;
    case "four-act":
      document.getElementById("structure-table").innerHTML = four_act(words, target);
      break;
    case "save-the-cat":
      document.getElementById("structure-table").innerHTML = save_the_cat(words, target);
      break;
  }
}