importScripts('https://unpkg.com/compromise@latest/builds/compromise.min.js');
importScripts('https://unpkg.com/compromise-syllables@0.0.6/builds/compromise-syllables.min.js');
self.nlp.extend(self.compromiseSyllables);

self.onmessage = function(e) {
    let doc = self.nlp.tokenize(e.data.text);
  
    let words = doc.wordCount();
  
    let characters = doc.termList().map(x => x.text).reduce((a,b) => a + b, "").length;
  
    let syllableList = doc.terms().syllables().map(x => x.syllables);
    let syllables = syllableList.map(x => x.length).reduce((a,b) => a + b, 0);
    let hardWords = (syllableList.filter(x => x.length > 2).length / words) * 100;
  
    hardWords = Math.min(Math.max(hardWords, 0), 1);
  
    // let sentences = doc.sentences().length;
    let sentences = doc.length;
  
    let ari = 4.71 * (characters / words) + 0.5 * (words / sentences) - 21.43;
    let fkr = 0.39 * (words / sentences) + 11.8 * (syllables / words) - 15.59;
    let gunning = 0.4 * ((words / sentences) + hardWords);

    self.postMessage({"characters": characters.toLocaleString(), "words": words.toLocaleString(), "sentences": sentences.toLocaleString(),
    "ari": ari, "fkr": fkr, "gunning": gunning});
};