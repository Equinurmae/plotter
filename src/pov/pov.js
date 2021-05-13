// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

import {ma} from 'moving-averages';

const d3 = require("d3");

// global variables

var pronoun_data = {"1st": 0, "2nd": 0, "3rd": 0};
var entities = [];
var lineChartData = [];

// web workers

var messageQueue = [];

const worker = new Worker("pov_worker.js");

worker.onmessage = function(e) {
  document.getElementById("debug").innerHTML = "Message received.";

  // update total counts and pie chart
  updatePronounData(e.data.pronouns);
  lineChartData.push(e.data.pronouns);

  entities = entities.concat(e.data.entities);

  redrawPieChart(messageQueue.length == 0);

  // check if messages left in queue
  if(messageQueue.length > 0) {  
    worker.postMessage({"text": messageQueue.pop()});
  } else {
    // update front end
    let dictionary = createDictionary(entities);
    addInfo(Object.entries(pronoun_data).sort((a,b) => a[1] - b[1])[2][0], dictionary);

    // draw line chart
    lineChartData.reverse();
    drawLineChart("1st");
  }
};

// creates a dictionary of keys to counts, sorted by most common key
function createDictionary(array)
{
    let entities = {};

    array.forEach((entity) => {
      if(entity in entities) {
        entities[entity] += 1;
      } else {
        entities[entity] = 1;
      }
    });

    return Object.entries(entities).sort((a,b) => b[1] - a[1]);
}

// global bar chart variables, so the chart can be updated in helper functions

var margin = {top: 30, right: 30, bottom: 30, left: 30}
, width = window.innerWidth - margin.left - margin.right
, height = window.innerWidth - margin.top - margin.bottom;

var radius = Math.min(width, height) / 2 - margin.top;

var svg = d3.select("#pie_chart_vis")
  .append("svg")
  .attr("width", width + margin.left + margin.right)
  .attr("height", height + margin.top + margin.bottom)
  .append("g")
  .attr("transform",
        "translate(" + ((width / 2) + margin.left) + "," + ((height / 2) + margin.top) + ")");

var colourScale = d3.scaleOrdinal(d3.schemeSet3);

colourScale.domain(Object.keys(pronoun_data));

var pie = d3.pie()
  .value((d) => d.value);

var arcGenerator = d3.arc()
  .innerRadius(0)
  .outerRadius(radius);

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.

    document.getElementById("refresh").onclick = refresh;
    document.getElementById("line_key").onchange = onLineKeyChange;

    refresh();
    drawPieChart();

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

// function called on pronoun dropdown change
function onLineKeyChange() {
  redrawLineChart(document.getElementById("line_key").value);
}

// function to update the spinners
function loading() {
  document.getElementById("summary").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("entities").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("line_chart_vis").innerHTML = `<div class="ms-Spinner"></div>`;

  var SpinnerElements = document.querySelectorAll(".ms-Spinner");
  for (var i = 0; i < SpinnerElements.length; i++) {
    new fabric['Spinner'](SpinnerElements[i]);
  }
}

// function to draw the message bars
function addInfo(person, dictionary) {
  document.getElementById("summary").innerHTML = `
  <div class="ms-MessageBar ms-MessageBar--success">
    <div class="ms-MessageBar-content">
      <div class="ms-MessageBar-icon">
        <i class="ms-Icon ms-Icon--Completed"></i>
      </div>
      <div class="ms-MessageBar-text">
        The selected text is most likely in <span style="font-weight: 700">` + person + ` person</span>, with <span style="font-weight: 700">` + dictionary[0][0] + 
        `</span> as the perspective character.
      </div>
    </div>
  </div>`;

  document.getElementById("entities").innerHTML = `
  <div class="ms-MessageBar">
    <div class="ms-MessageBar-content">
      <div class="ms-MessageBar-icon">
        <i class="ms-Icon ms-Icon--Info"></i>
      </div>
      <div class="ms-MessageBar-text">
        The most common entities in this text are <span style="font-weight: 700">` + dictionary[0][0] + `</span>, <span style="font-weight: 700">`
         + dictionary[1][0] + `</span> and <span style="font-weight: 700">` + dictionary[2][0] +`</span>.
      </div>
    </div>
  </div>`;
}

// main function
function refresh() {
  // reset all data
  loading();
  resetPronounData();

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

// function to draw the pie chart
function drawPieChart() {
  // get pie chart data
  var data_ready = pie(Object.entries(pronoun_data).map(function(x) {return {"key": x[0], "value": x[1]}; }));

  // draw the slices
  svg
    .selectAll('mySlices')
    .data(data_ready)
    .enter()
    .append('path')
    .attr("class", "arc")
    .attr('d', arcGenerator)
    .attr('fill', (d) => colourScale(d.data.key))
    .attr("stroke", "white")
    .style("stroke-width", "2px");

  // draw the labels
  svg
    .selectAll('mySlices')
    .data(data_ready)
    .enter()
    .append('text')
    .attr("class", "label")
    .text(function(d){ return d.data.key})
    .attr("transform", function(d) { return "translate(" + arcGenerator.centroid(d) + ")";  })
    .style("text-anchor", "middle")
    .style("font-size", 17)
    .style("opacity", d => d.data.value > 0 ? 100 : 0);
}

// function to redraw the pie chart
function redrawPieChart(animate) {
  // reset the chart data
  var data_ready = pie(Object.entries(pronoun_data).map(function(x) {return {"key": x[0], "value": x[1]}; }));

  // redraw the slices
  svg
    .selectAll('.arc')
    .data(data_ready)
    .transition().duration(animate ? 800 : 1)
    .attr('d', arcGenerator)
    .attr('fill', (d) => colourScale(d.data.key))
    .attr("stroke", "white")
    .style("stroke-width", "2px");

  // redraw the labels
  svg
    .selectAll('.label')
    .data(data_ready)
    .transition().duration(animate ? 800 : 10)
    .text(function(d){ return d.data.key})
    .attr("transform", function(d) { return "translate(" + arcGenerator.centroid(d) + ")";  })
    .style("text-anchor", "middle")
    .style("font-size", 17)
    .style("opacity", d => d.data.value > 0 ? 100 : 0);
}

// function to update the pronoun data
function updatePronounData(newData) {
  pronoun_data["1st"] += newData["1st"];
  pronoun_data["2nd"] += newData["2nd"];
  pronoun_data["3rd"] += newData["3rd"];
}

// function to reset the chart data
function resetPronounData() {
  pronoun_data = {"1st": 0, "2nd": 0, "3rd": 0};
  lineChartData = [];
}

// function to calculate the moving averages
function movingAverage(data, width) {
  return Array.from(ma(data, width));
}

// function to draw the line chart
function drawLineChart(key) {
  document.getElementById("line_chart_vis").innerHTML = "";

  var line_margin = {top: 50, right: 50, bottom: 50, left: 60}
    , line_width = window.innerWidth - line_margin.left - line_margin.right
    , line_height = window.innerHeight - line_margin.top - line_margin.bottom;

  // get the chart data
  var wordCounts = lineChartData.map(d => d["1st"] + d["2nd"] + d["3rd"]);
  var densities = lineChartData.map((d, i) => wordCounts[i] > 0 ? (d[key] / wordCounts[i]) : 0);
  var averages = movingAverage(densities, Math.ceil(densities.length * 0.15));

  // create the chart scales
  var xScaleLine = d3.scaleLinear()
    .domain([0, d3.max(densities)])
    .range([0, line_width]);

  var yScaleLine = d3.scaleLinear()
      .domain([lineChartData.length-1,0])
      .range([line_height, 0]);

  // create the line generator
  var line = d3.line()
    .x(function(d) { return xScaleLine(d); })
    .y(function(d, i) { return yScaleLine(i); })
    .curve(d3.curveBasis)
    .defined(d => !isNaN(d));

  // create the svg
  d3.select("#line_chart_vis").select("svg").remove();

  var line_svg = d3.select("#line_chart_vis").append("svg")
    .attr("width", line_width + line_margin.left + line_margin.right)
    .attr("height", line_height + line_margin.top + line_margin.bottom)
    .append("g")
      .attr("transform", "translate(" + line_margin.left + "," + line_margin.top + ")");

  // draw the y axis
  line_svg.append("g")
      .attr("class", "y-axis-line")
      .call(d3.axisLeft(yScaleLine));

  // draw the lines
  line_svg.append("path")
      .datum(densities)
      .attr("class", "line")
      .attr("stroke", "#0084ff")
      .attr("fill", "none")
      .attr("d", line);

  line_svg.append("path")
    .datum(averages)
    .attr("class", "lineAverage")
    .attr("d", line);

  // draw the x axis
  line_svg.append("g")
    .attr("class", "x-axis-line")
    .attr("transform", "translate(0,0)")
    .call(d3.axisTop(xScaleLine).ticks(5));

  line_svg.append("text")             
  .attr("transform",
        "translate(" + (line_width/2) + " ," + 
                       (- 30) + ")")
  .style("text-anchor", "middle")
  .text("Density");

  line_svg.append("text")
      .attr("transform", "rotate(-90)")
      .attr("y", 0 - line_margin.left + 10)
      .attr("x",0 - (line_height / 2))
      .attr("dy", "1em")
      .style("text-anchor", "middle")
      .text("Paragraph #");
}

// function to redraw the line chart
function redrawLineChart(key) {
  // update the chart data
  var wordCounts = lineChartData.map(d => d["1st"] + d["2nd"] + d["3rd"]);
  var densities = lineChartData.map((d, i) => wordCounts[i] > 0 ? (d[key] / wordCounts[i]) : 0);
  var averages = movingAverage(densities, Math.ceil(densities.length * 0.15));

  var svg = d3.select("#line_chart_vis").select("svg");
  
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

  // redraw the lines
  d3.selectAll(".line")
      .datum(densities)
      .transition().duration(800)
      .attr("d", line);

  d3.selectAll(".lineAverage")
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