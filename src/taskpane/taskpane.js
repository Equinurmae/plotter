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

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = refresh;

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
  Word.run(function (context) {
    loading();

      let paragraphs = context.document.body.paragraphs;
      paragraphs.load("text");

      let body = context.document.body;
      body.load("text");

      return context.sync()
        .then(function() {
          let doc = nlp(body.text);

          let words = doc.wordCount();

          let characters = doc.termList().map(x => x.text).reduce((a,b) => a + b, "").length;

          let syllableList = doc.terms().syllables();
          let syllables = syllableList.flatMap(x => x.syllables).length;
          let hardWords = (syllableList.map(x => x.syllables).filter(x => x.length > 2).length / words) * 100;

          hardWords = Math.min(Math.max(hardWords, 0), 1);

          let sentences = doc.sentences().length;

          let ari = 4.71 * (characters / words) + 0.5 * (words / sentences) - 21.43;
          let fkr = 0.39 * (words / sentences) + 11.8 * (syllables / words) - 15.59;
          let gunning = 0.4 * ((words / sentences) + hardWords);

          document.getElementById("character-count").innerHTML = characters.toLocaleString();
          document.getElementById("word-count").innerHTML = words.toLocaleString();
          document.getElementById("sentence-count").innerHTML = sentences.toLocaleString();
          document.getElementById("paragraph-count").innerHTML = paragraphs.items.length.toLocaleString();

          document.getElementById("ari").innerHTML = ari.toFixed(2);
          document.getElementById("fkr").innerHTML = fkr.toFixed(2);
          document.getElementById("gunning").innerHTML = gunning.toFixed(2);

          document.getElementById("ari-grade").innerHTML = gradeToAge(Math.round(ari));
          document.getElementById("fkr-grade").innerHTML = gradeToAge(Math.round(fkr));
          document.getElementById("gunning-grade").innerHTML = gradeToAge(Math.round(gunning));
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

function findSyllables(word) {
    word = word.toLowerCase();                                     
    word = word.replace(/(?:[^laeiouy]|ed|[^laeiouy]e)$/, '');   
    word = word.replace(/^y/, '');                                 
    //return word.match(/[aeiouy]{1,2}/g).length;   
    var syl = word.match(/[aeiouy]{1,2}/g);
    if(syl)
    {
        return syl.length;
    }
    else return 1;
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