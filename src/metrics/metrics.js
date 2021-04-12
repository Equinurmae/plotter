// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import nlp from "compromise";
import syllables from "compromise-syllables";
import sentences from "compromise-sentences"
nlp.extend(syllables);
nlp.extend(sentences);

/* global document, Office, Word */

var messageQueue = [];

var data = {characters: 0, words: 0, sentences: 0, syllables: 0, hardWords: 0};

const worker = new Worker("metrics_worker.js");

worker.onmessage = function(e) {
  document.getElementById("debug").innerHTML = "Message received.";

  data.characters += e.data.characters;
  data.words += e.data.words;
  data.sentences += e.data.sentences;
  data.syllables += e.data.syllables;
  data.hardWords += e.data.hardWords;

  if(messageQueue.length > 0) { 
    worker.postMessage({"text": messageQueue.pop()});
  } else {
    data.hardWords = (data.hardWords / data.words) * 100;

    data.hardWords = Math.min(Math.max(data.hardWords, 0), 1);

    let ari = 4.71 * (data.characters / data.words) + 0.5 * (data.words / data.sentences) - 21.43;
    let fkr = 0.39 * (data.words / data.sentences) + 11.8 * (data.syllables / data.words) - 15.59;
    let gunning = 0.4 * ((data.words / data.sentences) + data.hardWords);

    document.getElementById("character-count").innerHTML = data.characters.toLocaleString();
    document.getElementById("word-count").innerHTML = data.words.toLocaleString();
    document.getElementById("sentence-count").innerHTML = data.sentences.toLocaleString();
  
    document.getElementById("ari").innerHTML = ari.toFixed(2);
    document.getElementById("fkr").innerHTML = fkr.toFixed(2);
    document.getElementById("gunning").innerHTML = gunning.toFixed(2);
  
    document.getElementById("ari-grade").innerHTML = gradeToAge(Math.round(ari));
    document.getElementById("fkr-grade").innerHTML = gradeToAge(Math.round(fkr));
    document.getElementById("gunning-grade").innerHTML = gradeToAge(Math.round(gunning));
  }
};

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("refresh").onclick = refresh;

    refresh();

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function loading() {
  document.getElementById("character-count").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("word-count").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("sentence-count").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("paragraph-count").innerHTML = `<div class="ms-Spinner"></div>`;

  document.getElementById("ari").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("fkr").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("gunning").innerHTML = `<div class="ms-Spinner"></div>`;

  document.getElementById("ari-grade").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("fkr-grade").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("gunning-grade").innerHTML = `<div class="ms-Spinner"></div>`;

  var SpinnerElements = document.querySelectorAll(".ms-Spinner");
  for (var i = 0; i < SpinnerElements.length; i++) {
    new fabric['Spinner'](SpinnerElements[i]);
  }
}

function refresh() {
  loading();

  data = {characters: 0, words: 0, sentences: 0, syllables: 0, hardWords: 0};

  Word.run(function (context) {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");

    var selection = context.document.getSelection();

    selection.load("text");

    selection.paragraphs.load("text");

    return context.sync()
      .then(function() {
        document.getElementById("debug").innerHTML = "Message sending...";
        
        if(selection.text.length == 0) {
          document.getElementById("paragraph-count").innerHTML = paragraphs.items.length.toLocaleString();
          messageQueue = paragraphs.items.map(paragraph => paragraph.text);
        } else {
          document.getElementById("paragraph-count").innerHTML = selection.paragraphs.items.length.toLocaleString();
          let results = selection.paragraphs.items.map(paragraph => paragraph.text);

          if(results.length > 1) {
            let wholeText = results.join('\r');
            let match = new RegExp('(.*)' + selection.text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(.*)', 'g').exec(wholeText);

            if(match != null) {
              let firstParagraphMatch = new RegExp(match[1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(.*)', 'g').exec(results[0]);
              results[0] = firstParagraphMatch[1];

              let lastParagraphMatch = new RegExp('(.*)' + match[2].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g').exec(results[results.length-1]);
              results[results-1] = lastParagraphMatch[1];
            }
          } else {
            results[0] = selection.text;
          }

          messageQueue = results;
        }
        
        worker.postMessage({"text": messageQueue.pop()});
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

function gradeToAge(grade) {
    switch(grade) {
        case 0:
        return "n/a";
        case 1:
        return "5-6 y/o";
        case 2:
        return "6-7 y/o";
        case 3:
        return "7-9 y/o";
        case 4:
        return "9-10 y/o";
        case 5:
        return "10-11 y/o";
        case 6:
        return "11-12 y/o";
        case 7:
        return "12-13 y/o";
        case 8:
        return "13-14 y/o";
        case 9:
        return "14-15 y/o";
        case 10:
        return "15-16 y/o";
        case 11:
        return "16-17 y/o";
        case 12:
        return "17-18 y/o";
        case 13:
        return "18-24 y/o";
        case 14:
        return "24+ y/o";
        default:
        return "n/a";
    }
}