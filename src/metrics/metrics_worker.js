importScripts('https://unpkg.com/compromise@latest/builds/compromise.min.js');
importScripts('https://unpkg.com/compromise-syllables@0.0.6/builds/compromise-syllables.min.js');
self.nlp.extend(self.compromiseSyllables);

// tokenise and calculate counts
self.onmessage = function(e) {
    let doc = self.nlp.tokenize(e.data.text);
  
    let words = doc.wordCount();
  
    let characters = doc.termList().map(x => x.text).reduce((a,b) => a + b, "").length;
  
    let syllableList = doc.terms().syllables().map(x => x.syllables);
    let syllables = syllableList.map(x => x.length).reduce((a,b) => a + b, 0);
    let hardWords = syllableList.filter(x => x.length > 2).length;

    let sentences = doc.length;

    self.postMessage({"characters": characters, "words": words, "sentences": sentences,
    "syllables": syllables, "hardWords": hardWords});
};