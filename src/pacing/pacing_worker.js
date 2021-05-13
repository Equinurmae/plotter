importScripts('https://unpkg.com/compromise@latest/builds/compromise.min.js');
importScripts('https://unpkg.com/compromise-syllables@0.0.6/builds/compromise-syllables.min.js');
self.nlp.extend(self.compromiseSyllables);

// calculate the readability
self.onmessage = function(e) {
    let doc = self.nlp.tokenize(e.data.text);
  
    let words = doc.wordCount();

    let sentences = doc.length;
   
    let syllableList = doc.terms().syllables().map(x => x.syllables);
    let hardWords = syllableList.filter(x => x.length > 2).length;
    hardWords = (hardWords / words) * 100;
    hardWords = Math.min(Math.max(hardWords, 0), 1);

    let gunning = 0.4 * ((words / sentences) + hardWords);

    self.postMessage({"words": words, "readability": gunning});
};