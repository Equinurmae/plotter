// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

import {
  ma, dma, ema, sma, wma
} from 'moving-averages';

const d3 = require("d3");

var options = {"readability": true, "words": true, "average": true, "z-scores": false, "detail": 100};

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("detail").value = options.detail;

    document.getElementById("insert-paragraph").onclick = refresh;
    document.getElementById("readability").onchange = onReadability;
    document.getElementById("words").onchange = onWords;
    document.getElementById("average").onchange = onAverage;
    document.getElementById("detail").oninput = onDetail;
    document.getElementById("z-scores").onchange = onZ;

    refresh(true);

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function onReadability() {
  options.readability = !options.readability;
  refresh();
}

function onWords() {
  options.words = !options.words;
  refresh();
}

function onAverage() {
  options.average = !options.average;
  refresh();
}

function onZ() {
  options["z-scores"] = !options["z-scores"];
  refresh();
}

function onDetail() {
  options.detail = document.getElementById("detail").value;
  document.getElementById("detail_label").innerHTML = options.detail;
  refresh();
}

function refresh(first = false) {
  Word.run(function (context) {
      let paragraphs = context.document.body.paragraphs;
      paragraphs.load("text");

      return context.sync()
        .then(function() {
          document.getElementById("paragraph-count").innerHTML = paragraphs.items.length.toLocaleString();
          if(first) document.getElementById("detail").value = Math.ceil(paragraphs.items.length * 0.75);

          draw_chart(paragraphs);
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

function getWordCount(text) {
  let strip_punctuation = text.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()"?“”]/g," ");
  let words = strip_punctuation.trim().split(/\s+/g);
  return words.length;
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

function getReadability(text) {
  let strip_punctuation = text.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()"?“”]/g," ");
  let words = strip_punctuation.trim().split(/\s+/g);

  let syllableList = words.map(word => findSyllables(word));
  let hardWords = (syllableList.filter(x => x > 2).length / words.length) * 100;

  let sentences = text.match(/\w[.?!](\s|$|”)/g);

  hardWords = Math.min(Math.max(hardWords, 0), 1);

  return 0.4 * ((words.length / (sentences == null ? 1 : sentences.length)) + hardWords);
}

function movingAverage(data, width) {
  return Array.from(ma(data, width));
}

function exponentialMovingAverage(data, width) {
  return Array.from(sma(data, width));
}

function calculateCI(point, mean, deviation, n) {
  let z = (point - mean) / deviation;
  return {"upper": mean + (z * (deviation / Math.sqrt(n)) ), "lower": mean - (z * (deviation / Math.sqrt(n)) )};
}

function z(point, mean, deviation) {
  return Math.abs((point - mean) / deviation);
}

function draw_chart(paragraphs) {
  var margin = {top: 50, right: 50, bottom: 50, left: 50}
  , width = window.innerWidth - margin.left - margin.right // Use the window's width 
  , height = window.innerHeight - margin.top - margin.bottom; // Use the window's height

  let data = paragraphs.items.map(paragraph => getReadability(paragraph.text));

  let words = paragraphs.items.map(paragraph => getWordCount(paragraph.text));

  document.getElementById("detail").min = 2;
  document.getElementById("detail").max = data.length - 2;
  let averages = movingAverage(data, options.detail);
  let emas = exponentialMovingAverage(data, options.detail);

  let mean = d3.mean(data);
  let deviation = d3.deviation(data);

  let zs = data.map(d => z(d, mean, deviation));

  // let confidence_interval = data.map(d => calculateCI(d, d3.mean(data), d3.deviation(data), data.length));

// 5. X scale will use the index of our data
var xScale = d3.scaleLinear()
    .domain([0, d3.max(data)]) // input
    .range([0, width]); // output

    // 5. X scale will use the index of our data
var xWordScale = d3.scaleLinear()
.domain([0, d3.max(words)]) // input
.range([0, width]); // output

    // 5. X scale will use the index of our data
    var xZScale = d3.scaleLinear()
    .domain([0, d3.max(zs)]) // input
    .range([0, width]); // output

// 6. Y scale will use the randomly generate number 
var yScale = d3.scaleLinear()
    .domain([data.length-1,0]) // input 
    .range([height, 0]); // output 

// 7. d3's line generator
var line = d3.line()
    .x(function(d) { return xScale(d.y); }) // set the x values for the line generator
    .y(function(d, i) { return yScale(i); }) // set the y values for the line generator 
    .curve(d3.curveBasis)
    .defined(d => d.y != undefined); // apply smoothing to the line

// var area = d3.area()
//   .x0(function(d) { return xScale(d.lower); }) // set the x values for the line generator
//   .x1(function(d) { return xScale(d.upper); })
//   .y(function(d, i) { return yScale(i); }); // apply smoothing to the line

var wordLine = d3.line()
  .x(function(d) { return xWordScale(d.y); }) // set the x values for the line generator
  .y(function(d, i) { return yScale(i); }) // set the y values for the line generator 
  .curve(d3.curveMonotoneX)
  .defined(d => d.y != undefined); // apply smoothing to the line

  var zLine = d3.line()
  .x(function(d) { return xZScale(d.y); }) // set the x values for the line generator
  .y(function(d, i) { return yScale(i); }) // set the y values for the line generator 
  .curve(d3.curveMonotoneX)
  .defined(d => d.y != undefined); // apply smoothing to the line


// 8. An array of objects of length N. Each object has key -> value pair, the key being "y" and the value is a random number
var dataset = data.map(function(d) { return {"y": d } });

var wordDataset = words.map(function(d) { return {"y": d } });

var zDataset = zs.map(function(d) { return {"y": d } });

var averageDataset = averages.map(function(d) { return {"y": d } });

var emaDataset = emas.map(function(d) { return {"y": d } });

// 1. Add the SVG to the page and employ #2
d3.select("svg").remove();
var svg = d3.select("#pacing_vis").append("svg")
    .attr("width", width + margin.left + margin.right)
    .attr("height", height + margin.top + margin.bottom)
  .append("g")
    .attr("transform", "translate(" + margin.left + "," + margin.top + ")");

// 3. Call the x axis in a group tag
// svg.append("g")
//     .attr("class", "x axis")
//     .attr("transform", "translate(0,0)")
//     .call(d3.axisTop(xScale)); // Create an axis component with d3.axisBottom

// 4. Call the y axis in a group tag
svg.append("g")
    .attr("class", "y axis")
    .call(d3.axisLeft(yScale)); // Create an axis component with d3.axisLeft

if(options.words) {
  // 9. Append the path, bind the data, and call the line generator 
svg.append("path")
.datum(wordDataset) // 10. Binds data to the line 
.attr("class", "wordLine") // Assign a class for styling 
.attr("d", wordLine); // 11. Calls the line generator 
}

if(options.readability) {
// 9. Append the path, bind the data, and call the line generator 
svg.append("path")
    .datum(dataset) // 10. Binds data to the line 
    .attr("class", "line") // Assign a class for styling 
    .attr("d", line); // 11. Calls the line generator 
}

if(options["z-scores"]) {
  svg.append("path")
  .datum(zDataset) // 10. Binds data to the line 
  .attr("class", "lineZ") // Assign a class for styling 
  .attr("d", zLine); // 11. Calls the line generator 
}

  if(options.average) {
  // 9. Append the path, bind the data, and call the line generator 
  svg.append("path")
  .datum(averageDataset) // 10. Binds data to the line 
  .attr("class", "lineAverage") // Assign a class for styling 
  .attr("d", line); // 11. Calls the line generator 

  svg.append("path")
  .datum(emaDataset) // 10. Binds data to the line 
  .attr("class", "ema") // Assign a class for styling 
  .attr("d", line); // 11. Calls the line generator 
  }


// svg.append("path")
//       .datum(confidence_interval)
//       .attr("fill", "green")
//       .attr("stroke", "none")
//       .attr("d", area);
}