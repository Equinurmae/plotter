// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

import {ma} from 'moving-averages';

const d3 = require("d3");

// global variables

var barChartData = [
  {"name": "Adjectives", "count": 0},
  {"name": "Adverbs", "count": 0},
  {"name": "Conjunctions", "count": 0},
  {"name": "Determiners", "count": 0},
  {"name": "Nouns", "count": 0},
  {"name": "Pronouns", "count": 0},
  {"name": "Proper Nouns", "count": 0},
  {"name": "Prepositions", "count": 0},
  {"name": "Verbs", "count": 0}
];

var lineChartData = [];

var totalActive = 0;
var totalPassive = 0;

// web workers

var messageQueue = [];

const worker = new Worker("pos_worker.js");

worker.onmessage = function(e) {
  document.getElementById("debug").innerHTML = "Message received.";

  // update total counts and bar chart
  updateBarChartData(e.data.pos);
  lineChartData.push(e.data.pos);
  totalActive += e.data.active;
  totalPassive += e.data.passive;

  redrawBarChart(messageQueue.length == 0);

  // check if messages left in queue
  if(messageQueue.length > 0) {
    worker.postMessage({"text": messageQueue.pop()});
  } else {
    // update message bars
    getActivePassiveInfo(e.data.active, e.data.passive);

    document.getElementById("notifications").innerHTML = "";
    let words = barChartData.map(x => x.count).reduce((a,b) => a + b, 0);
    getWordTypeInfo(0, words, 0.1, "adjectives");
    getWordTypeInfo(1, words, 0.05, "adverbs");

    // update line chart
    lineChartData.reverse();
    drawLineChart(0);
  }
};

// global bar chart variables, so the chart can be updated in helper functions

document.getElementById("pos_vis").innerHTML = "";

var margin = {top: 30, right: 30, bottom: 50, left: 80}
, width = window.innerWidth - margin.left - margin.right
, height = (5 * 50) - margin.top - margin.bottom;

var svg = d3.select("#pos_vis")
  .append("svg")
  .attr("width", width + margin.left + margin.right)
  .attr("height", height + margin.top + margin.bottom)
  .append("g")
  .attr("transform",
        "translate(" + margin.left + "," + margin.top + ")");

var xScale = d3.scaleLinear()
  .domain([0, 100])
  .range([ 0, width]);

var yScale = d3.scaleBand()
  .range([ 0, height ])
  .domain(barChartData.map(function(d) { return d.name; }))
  .padding(.1);

var colourScale = d3.scaleSequential()
  .domain([0,d3.max(barChartData.map(d => d.count))])
  .interpolator(d3.interpolateYlGnBu);

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("refresh").onclick = refresh;
    document.getElementById("line_index").onchange = onLineIndexChange;

    drawBarChart();
    refresh();

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

// function to display the spinners
function loading() {
  document.getElementById("active-passive").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("notifications").innerHTML = `<br><div class="ms-Spinner"></div><br>`;
  document.getElementById("line_chart_vis").innerHTML = `<div class="ms-Spinner"></div>`;

  var SpinnerElements = document.querySelectorAll(".ms-Spinner");
  for (var i = 0; i < SpinnerElements.length; i++) {
    new fabric['Spinner'](SpinnerElements[i]);
  }
}

// main function
function refresh() {
  // reset all data
  loading();
  resetBarChartData();

  Word.run(function (context) {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");

    var selection = context.document.getSelection();
    selection.load("text");

    selection.paragraphs.load("text");

    return context.sync()
      .then(function() {
        document.getElementById("debug").innerHTML = "Message sending...";
        
        // get selection and split into paragraphs
        if(selection.text.length == 0) {
          // no selection, so use body text
          messageQueue = paragraphs.items.map(paragraph => paragraph.text);
        } else {
          // use selection
          let results = selection.paragraphs.items.map(paragraph => paragraph.text);

          if(results.length > 1) {
            // one or more paragraphs selected
            // use regex to find the intersections
            let wholeText = results.join('\r');
            let match = new RegExp('(.*)' + selection.text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(.*)', 'g').exec(wholeText);

            if(match != null) {
              let firstParagraphMatch = new RegExp(match[1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(.*)', 'g').exec(results[0]);
              results[0] = firstParagraphMatch[1];

              let lastParagraphMatch = new RegExp('(.*)' + match[2].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g').exec(results[results.length-1]);
              results[results-1] = lastParagraphMatch[1];
            }
          } else {
            // one or fewer paragraphs selected
            results[0] = selection.text;
          }

          // update message queue
          messageQueue = results;
        }

        // start web worker processing
        
        drawBarChart();

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

// function triggered on POS dropdown change
function onLineIndexChange() {
  redrawLineChart(parseInt(document.getElementById("line_index").value));
}

// function to reset POS data
function resetBarChartData() {
  barChartData = [
    {"name": "Adjectives", "count": 0},
    {"name": "Adverbs", "count": 0},
    {"name": "Conjunctions", "count": 0},
    {"name": "Determiners", "count": 0},
    {"name": "Nouns", "count": 0},
    {"name": "Pronouns", "count": 0},
    {"name": "Proper Nouns", "count": 0},
    {"name": "Prepositions", "count": 0},
    {"name": "Verbs", "count": 0}
  ];

  lineChartData = [];

  totalActive = 0;
  totalPassive = 0;
}

// function to update bar chart data
function updateBarChartData(newData) {
  for(let i = 0; i < barChartData.length; i++) {
    barChartData[i].count += newData[i].count;
  }
}

// function to display active/passive message bar
function getActivePassiveInfo() {
  let sentences = totalActive + totalPassive;

  let active = ((totalActive / sentences) * 100).toFixed(2);
  let passive = ((totalPassive / sentences) * 100).toFixed(2);

  document.getElementById("active-passive").innerHTML = `
    <div class="ms-MessageBar ms-MessageBar--` + (active > passive ? `success` : `error`) + `">
      <div class="ms-MessageBar-content">
        <div class="ms-MessageBar-icon">
          <i class="ms-Icon ms-Icon--` + (active > passive ? `Completed` : `ErrorBadge`) + `"></i>
        </div>
        <div class="ms-MessageBar-text">` + (active > passive ? 
            `<span style="font-weight: 700">` + active + `% Active</span> | ` + passive + `% Passive`
            : active + `% Active | <span style="font-weight: 700">` + passive + `% Passive</span>`) + `
        </div>
      </div>
    </div>`;
}

// function to display word type message bar
function getWordTypeInfo(index, words, tolerance, name) {
  let percentage = barChartData[index].count / words;

  document.getElementById("notifications").innerHTML += `
    <div class="ms-MessageBar ms-MessageBar--` + (percentage <= tolerance ? `success` : `warning`) + `">
      <div class="ms-MessageBar-content">
        <div class="ms-MessageBar-icon">
          <i class="ms-Icon ms-Icon--` + (percentage <= tolerance ? `Completed` : `Info`) + `"></i>
        </div>
        <div class="ms-MessageBar-text">` + (percentage <= tolerance ?
          `That's a nice ratio of <span style="font-weight: 700">` + name + `</span> you've got there!` :
          `Woah! That's a lot of <span style="font-weight: 700">` + name + `</span> there!`) + `
        </div>
      </div>
    </div>`;
}

// function to draw the bar chart
function drawBarChart() {
  d3.select("svg").remove();

  margin = {top: 30, right: 40, bottom: 50, left: 70}
  , width = window.innerWidth - margin.left - margin.right
  , height = (5 * 50) - margin.top - margin.bottom;

  svg = d3.select("#pos_vis")
    .append("svg")
    .attr("width", width + margin.left + margin.right)
    .attr("height", height + margin.top + margin.bottom)
    .append("g")
    .attr("transform",
          "translate(" + margin.left + "," + margin.top + ")");

  // create the scales for the bar chart
  xScale = d3.scaleLinear()
    .domain([0, 100])
    .range([ 0, width]);

  yScale = d3.scaleBand()
    .range([ 0, height ])
    .domain(barChartData.map(function(d) { return d.name; }))
    .padding(.1);

  colourScale = d3.scaleSequential()
    .domain([0,d3.max(barChartData.map(d => d.count))])
    .interpolator(d3.interpolateYlGnBu);

  // draw the bars
  svg.selectAll("myRect")
    .data(barChartData)
    .enter()
    .append("rect")
    .attr("class", "bar")
    .attr("x", xScale(0) )
    .attr("y", function(d) { return yScale(d.name); })
    .attr("width", 0)
    .attr("height", yScale.bandwidth() )
    .attr("fill", d => colourScale(d.count));

  // animate the bars
  svg.selectAll("rect")
    .transition()
    .duration(800)
    .attr("width", function(d) { return xScale(d.count); })
    .delay(function(d,i){return(i*100);})

  // draw the axes
  svg.append("g")
    .attr("class", "x-axis")
    .attr("transform", "translate(0," + height + ")")
    .call(d3.axisBottom(xScale))
    .selectAll("text")
      .attr("transform", "translate(-10,0)rotate(-45)")
      .style("text-anchor", "end");

  svg.append("g")
    .call(d3.axisLeft(yScale));

  // draw the axes labels
  svg.append("text")             
  .attr("transform",
        "translate(" + (width/2) + " ," + 
                        (height + 40) + ")")
  .style("text-anchor", "middle")
  .text("Count");
}

function redrawBarChart(animate = true) {
  // remake the scales
  xScale = d3.scaleLinear()
    .domain([0, d3.max(barChartData.map(d => d.count))])
    .range([ 0, width]);

  colourScale = d3.scaleSequential()
  .domain([0,d3.max(barChartData.map(d => d.count))])
  .interpolator(d3.interpolateYlGnBu);

  // redraw the bars
  d3.selectAll(".bar")
      .data(barChartData)
      .transition().duration(animate ? 800 : 10)
      .attr("x", xScale(0) )
      .attr("y", function(d) { return yScale(d.name); })
      .attr("width", function(d) { return xScale(d.count); })
      .attr("fill", d => colourScale(d.count));
  
  // redraw the x axis with the new scale
  svg.selectAll(".x-axis").remove();

  svg.append("g")
    .attr("class", "x-axis")
    .attr("transform", "translate(0," + height + ")")
    .call(d3.axisBottom(xScale))
    .selectAll("text")
      .attr("transform", "translate(-10,0)rotate(-45)")
      .style("text-anchor", "end");
}

// function to calculate the moving averages
function movingAverage(data, width) {
  return Array.from(ma(data, width));
}

// function to draw the line chart
function drawLineChart(index) {
  document.getElementById("line_chart_vis").innerHTML = "";

  // get the data for the lines
  var wordCounts = lineChartData.map(d => d.map(x => x.count).reduce((a, b) => a + b, 0));
  var densities = lineChartData.map((d, i) => wordCounts[i] > 0 ? (d[index].count / wordCounts[i]) : 0);
  var averages = movingAverage(densities, Math.ceil(lineChartData.length * 0.15));

  var margin = {top: 50, right: 50, bottom: 50, left: 60}
    , width = window.innerWidth - margin.left - margin.right
    , height = window.innerHeight - margin.top - margin.bottom;

  // create the scales for the line chart
  var xScaleLine = d3.scaleLinear()
    .domain([0, d3.max(densities)])
    .range([0, width]);

  var xScaleWordLine = d3.scaleLinear()
    .domain([0, d3.max(wordCounts)])
    .range([0, width]);

  var yScaleLine = d3.scaleLinear()
      .domain([densities.length-1,0])
      .range([height, 0]);

  // create the line generators
  var line = d3.line()
    .x(function(d) { return xScaleLine(d); })
    .y(function(d, i) { return yScaleLine(i); })
    .curve(d3.curveBasis)
    .defined(d => !isNaN(d));
  
  var wordLine = d3.line()
    .x(function(d) { return xScaleWordLine(d); })
    .y(function(d, i) { return yScaleLine(i); })
    .curve(d3.curveBasis)
    .defined(d => !isNaN(d));

  // draw the chart
  d3.select("#line_chart_vis").select("svg").remove();

  var svg = d3.select("#line_chart_vis").append("svg")
    .attr("width", width + margin.left + margin.right)
    .attr("height", height + margin.top + margin.bottom)
    .append("g")
      .attr("transform", "translate(" + margin.left + "," + margin.top + ")");

  // draw the y axis
  svg.append("g")
      .attr("class", "y-axis-line")
      .call(d3.axisLeft(yScaleLine));

  // draw the lines
  svg.append("path")
    .datum(wordCounts)
    .attr("class", "word-line")
    .attr("stroke", "#c9e5ff")
    .attr("fill", "none")
    .attr("stroke-dasharray", "5px")
    .attr("d", wordLine);

  svg.append("path")
      .datum(densities)
      .attr("class", "line")
      .attr("stroke", "#0084ff")
      .attr("fill", "none")
      .attr("d", line);

  svg.append("path")
    .datum(averages)
    .attr("class", "lineAverage")
    .attr("d", line);

  // draw the x axis
  svg.append("g")
    .attr("class", "x-axis-line")
    .attr("transform", "translate(0,0)")
    .call(d3.axisTop(xScaleLine));

  svg.append("text")             
  .attr("transform",
        "translate(" + (width/2) + " ," + 
                       (- 30) + ")")
  .style("text-anchor", "middle")
  .text("Density");

  svg.append("text")
      .attr("transform", "rotate(-90)")
      .attr("y", 0 - margin.left + 10)
      .attr("x",0 - (height / 2))
      .attr("dy", "1em")
      .style("text-anchor", "middle")
      .text("Paragraph #");
}

function redrawLineChart(index) {
  var svg = d3.select("#line_chart_vis").select("svg");

  // recalculate the line chart data
  var wordCounts = lineChartData.map(d => d.map(x => x.count).reduce((a, b) => a + b, 0));
  var densities = lineChartData.map((d, i) => wordCounts[i] > 0 ? (d[index].count / wordCounts[i]) : 0);
  var averages = movingAverage(densities, Math.ceil(densities.length * 0.15));

  var margin = {top: 50, right: 50, bottom: 50, left: 60}
    , width = window.innerWidth - margin.left - margin.right
    , height = window.innerHeight - margin.top - margin.bottom;

  // remake the scales
  var xScaleLine = d3.scaleLinear()
    .domain([0, d3.max(densities)])
    .range([0, width]);

  var yScaleLine = d3.scaleLinear()
    .domain([lineChartData.length-1,0])
    .range([height, 0]);

  // remake the line generator
  var line = d3.line()
    .x(function(d) { return xScaleLine(d); })
    .y(function(d, i) { return yScaleLine(i); })
    .curve(d3.curveBasis)
    .defined(d => !isNaN(d));

  // redraw the line
  d3.selectAll(".line")
      .datum(densities)
      .transition().duration(800)
      .attr("d", line);
  
  svg.selectAll(".lineAverage")
    .datum(averages)
    .transition().duration(800)
    .attr("d", line);
  
  // redraw the x axis with the new scale
  svg.selectAll(".x-axis-line").remove();

  svg.append("g")
    .attr("class", "x-axis-line")
    .attr("transform", "translate(" + margin.left + "," + margin.top + ")")
    .call(d3.axisTop(xScaleLine).ticks(5));
}