importScripts('https://unpkg.com/compromise@latest/builds/compromise.min.js');
importScripts('https://unpkg.com/compromise-sentences@0.2.0/builds/compromise-sentences.min.js');
self.nlp.extend(self.compromiseSentences);

self.onmessage = function(e) {
    let doc = self.nlp(e.data.text);

    self.postMessage(self.pos(doc));
};

function pos(doc) {
    let pronouns = doc.match("#Pronoun").length;
    let proper_nouns = doc.match("#ProperNoun").length;
    let verbs = doc.match("#Verb").length;
    let adjectives = doc.match("#Adjective").length;
    let adverbs = doc.match("#Adverb").length;
    let conjunctions = doc.match("#Conjunction").length;
    let prepositions = doc.match("#Preposition").length;
    let determiners = doc.match("#Determiner").length;

    let nouns = doc.match("#Noun").length - pronouns - proper_nouns;

    let sentences = doc.sentences();

    let passive = sentences.isPassive().out('array').length;
    let active = sentences.length - passive;

    return {"pos": [
        {"name": "Adjectives", "count": adjectives},
        {"name": "Adverbs", "count": adverbs},
        {"name": "Conjunctions", "count": conjunctions},
        {"name": "Determiners", "count": determiners},
        {"name": "Nouns", "count": nouns},
        {"name": "Pronouns", "count": pronouns},
        {"name": "Proper Nouns", "count": proper_nouns},
        {"name": "Prepositions", "count": prepositions},
        {"name": "Verbs", "count": verbs},
    ], "passive": passive, "active": active};
}