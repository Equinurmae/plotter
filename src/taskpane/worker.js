// import nlp from "compromise";
// import syllables from "compromise-syllables";
// import sentences from "compromise-sentences"
// nlp.extend(syllables);
// nlp.extend(sentences);

self.addEventListener('message', function(e) {
    // self.postMessage(calculate(e.text, e.paragraphs));
    self.postMessage({"characters": 0, "words": 0, "sentences": 0, "paragraphs": 0, "ari": 0, "fkr": 0, "gunning": 0});
});

// function calculate(text, paragraphs) {
//     let doc = nlp(text);
  
//     let words = doc.wordCount();
  
//     let characters = doc.termList().map(x => x.text).reduce((a,b) => a + b, "").length;
  
//     let syllableList = doc.terms().syllables();
//     let syllables = syllableList.flatMap(x => x.syllables).length;
//     let hardWords = (syllableList.map(x => x.syllables).filter(x => x.length > 2).length / words) * 100;
  
//     hardWords = Math.min(Math.max(hardWords, 0), 1);
  
//     let sentences = doc.sentences().length;
  
//     let ari = 4.71 * (characters / words) + 0.5 * (words / sentences) - 21.43;
//     let fkr = 0.39 * (words / sentences) + 11.8 * (syllables / words) - 15.59;
//     let gunning = 0.4 * ((words / sentences) + hardWords);

//     return {"characters": characters.toLocaleString(), "words": words.toLocaleString(), "sentences": sentences.toLocaleString(),
//             "paragraphs": paragraphs.items.length.toLocaleString(), "ari": ari, "fkr": fkr, "gunning": gunning};
//   }