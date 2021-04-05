importScripts('https://unpkg.com/compromise@latest/builds/compromise.min.js');

self.onmessage = function(e) {
    let doc = self.nlp(e.data.text);

    self.postMessage(self.count_pronouns(doc));
};

function count_pronouns(doc) {
    let first = doc.match("(I|me|my|mine|myself|we|us|our|ours|ourselves)").length;
    let second = doc.match("(you|your|yours|yourself|yourselves)").length;
    let third = doc.match("(he|him|his|himself|she|her|hers|herself|it|its|itself|they|them|their|theirs|themselves|themself)").length;
    let male = doc.match("(he|him|his|himself)").length;
    let female = doc.match("(she|her|hers|herself)").length;
    let neutral = doc.match("(it|its|itself|they|them|their|theirs|themselves|themself)").length;

    return {"pronouns" : {"1st": first, "2nd": second, "3rd": third, "male": male, "female": female, "neutral": neutral}};
}